#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VSTO Inclusion List 工具：把加载项清单加入/移出当前用户的「受信任清单」。

背景
----
用 |vstolocal 部署 VSTO 加载项时，Office 除了读
  HKCU\\Software\\Microsoft\\Office\\Outlook\\Addins\\<Name>  (Manifest / LoadBehavior)
还要求该 .vsto 出现在 **Inclusion List** 里：
  HKCU\\Software\\Microsoft\\VSTO\\Security\\Inclusion\\<GUID>
    Url       = 清单的完整 file:/// URL
    PublicKey = 签名证书公钥的 XML 形式 (<RSAKeyValue>…</RSAKeyValue>)

缺少 Inclusion 条目时，Office 会**静默拒绝加载**（功能区没有任何按钮，
事件 ID 45 也不会列出该加载项），非常难排查。

正式 API 是 Microsoft.VisualStudio.Tools.Applications.Deployment 里的
UserInclusionList.Add/Remove（需要加载 VSTO 运行时程序集）。本脚本用等价方式
直接写注册表：从 .vsto 里取出签名证书，解析出 RSA 公钥，写成同样的 XML 格式。

用法：
    python tools/vsto_inclusion.py add     <path-to-vsto>
    python tools/vsto_inclusion.py remove  <path-to-vsto>
    python tools/vsto_inclusion.py list
"""
import base64
import os
import re
import sys
import uuid
import winreg

# HKCU\Software\Microsoft\VSTO\Security\Inclusion
INCLUSION_KEY = r"Software\Microsoft\VSTO\Security\Inclusion"


def read_vsto_public_key(vsto_path):
    """从 .vsto 清单取出签名证书公钥，返回 VSTO 使用的 XML 字符串。

    流程：.vsto -> <X509Certificate> base64 DER -> SubjectPublicKeyInfo ->
          RSA (Modulus, Exponent) -> <RSAKeyValue>…</RSAKeyValue>

    实现说明：先用 openssl（Windows 10+ 自带）把 DER 证书转成 PKCS#1 的
    RSA 公钥（"-----BEGIN RSA PUBLIC KEY-----"），再解析那两个整数。
    这样避免了手写 ASN.1 解析容易在 TBSCertificate 字段上错位。
    没有 openssl 时回退到手写 ASN.1 解析。
    """
    with open(vsto_path, "rb") as f:
        data = f.read().decode("utf-8", "replace")

    m = re.search(r"<X509Certificate>(.*?)</X509Certificate>", data, re.S)
    if not m:
        raise SystemExit("ERROR: no <X509Certificate> found in %s" % vsto_path)
    der = base64.b64decode(re.sub(r"\s+", "", m.group(1)))

    modulus, exponent = _rsa_public_key_via_openssl(der)
    if modulus is None:
        modulus, exponent = _parse_rsa_public_key(der)

    # .NET 的 RSA.ExportParameters(false) 会去掉 modulus 的前导 0 字节（256 字节标准长度），
    # VSTO 写入 Inclusion List 时用的就是这种形式，必须去掉才能与 Office 生成的条目一致。
    modulus = modulus.lstrip(b"\x00") or b"\x00"

    b64mod = base64.b64encode(modulus).decode()
    b64exp = base64.b64encode(exponent).decode()
    return "<RSAKeyValue><Modulus>%s</Modulus><Exponent>%s</Exponent></RSAKeyValue>" % (b64mod, b64exp)


def _rsa_public_key_via_openssl(der):
    """用 openssl 把 DER 证书转成 PKCS#1 RSA 公钥并解析出 (modulus, exponent)。

    返回 (None, None) 表示不可用（调用方回退到内置解析）。
    """
    import shutil
    import subprocess
    import tempfile

    if not shutil.which("openssl"):
        return None, None
    tmpdir = tempfile.mkdtemp(prefix="vsto_inc_")
    cert_path = os.path.join(tmpdir, "cert.der")
    with open(cert_path, "wb") as f:
        f.write(der)
    try:
        r = subprocess.run(
            ["openssl", "x509", "-inform", "DER", "-in", cert_path,
             "-pubkey", "-noout"],
            capture_output=True, text=True)
        if r.returncode != 0:
            return None, None
        pem = r.stdout
        r2 = subprocess.run(
            ["openssl", "rsa", "-pubin", "-inform", "PEM", "-RSAPublicKey_out", "-outform", "DER"],
            input=pem.encode(), capture_output=True)
        if r2.returncode != 0:
            return None, None
        pk = r2.stdout
        # PKCS#1 RSAPublicKey ::= SEQUENCE { modulus INTEGER, publicExponent INTEGER }
        # openssl -outform DER 输出带外层 SEQUENCE 的 tag+length，故先跳过它
        _tag, ln, i = _read_tlv(pk, 0)
        rsa = pk[i:i + ln]
        _tag, mod_len, i = _read_tlv(rsa, 0)
        modulus = rsa[i:i + mod_len]
        i += mod_len
        _tag, exp_len, i = _read_tlv(rsa, i)
        exponent = rsa[i:i + exp_len]
        return modulus, exponent
    except Exception:
        return None, None
    finally:
        try:
            for fn in os.listdir(tmpdir):
                os.remove(os.path.join(tmpdir, fn))
            os.rmdir(tmpdir)
        except Exception:
            pass


def _parse_rsa_public_key(der):
    """从 X.509 DER 里取出 RSA 公钥的 (modulus, exponent) 字节。

    手工做最小 ASN.1 解析，避免引入 cryptography 依赖：
      Certificate ::= SEQUENCE { tbsCertificate, ... }
      TBSCertificate ::= SEQUENCE {
          [0] version, serialNumber, signature, issuer, validity, subject,
          subjectPublicKeyInfo, ... }
      SubjectPublicKeyInfo ::= SEQUENCE { algorithm, subjectPublicKey BIT STRING }
      RSAPublicKey ::= SEQUENCE { modulus INTEGER, publicExponent INTEGER }
    """
    # 剥掉 Certificate 外层 SEQUENCE
    cert = _der_seq(der)[1]
    tbs = _der_seq(cert)[1]

    # 在 TBSCertificate 里找 subjectPublicKeyInfo（跳过前面的可选/基本字段）
    idx = 0
    # version [0] 是 context tag 0xA0
    if tbs[idx] == 0xA0:
        _tag, _l, idx = _read_tlv(tbs, idx)
    # serialNumber INTEGER
    _tag, _l, idx = _read_tlv(tbs, idx)
    # signature SEQUENCE
    _tag, _l, idx = _read_tlv(tbs, idx)
    # issuer SEQUENCE
    _tag, _l, idx = _read_tlv(tbs, idx)
    # validity SEQUENCE
    _tag, _l, idx = _read_tlv(tbs, idx)
    # subject SEQUENCE
    _tag, _l, idx = _read_tlv(tbs, idx)
    # subjectPublicKeyInfo SEQUENCE
    _tag, spki_len, idx = _read_tlv(tbs, idx)
    spki = tbs[idx:idx + spki_len]

    # SPKI: algorithm SEQUENCE, subjectPublicKey BIT STRING
    _tag, _l, i = _read_tlv(spki, 0)
    bitstring = spki[i:]
    # 去掉 BIT STRING 的 unused-bits 字节
    pk = bitstring[1:]
    # RSAPublicKey ::= SEQUENCE { modulus, exponent }
    _tag, _l, i = _read_tlv(pk, 0)
    rsa = pk[i:]
    _tag, mod_len, i = _read_tlv(rsa, 0)
    modulus = rsa[i:i + mod_len]
    i += mod_len
    _tag, exp_len, i = _read_tlv(rsa, i)
    exponent = rsa[i:i + exp_len]
    return modulus, exponent


def _read_tlv(buf, i):
    """读一个 TLV，返回 (tag, length, next_index)。"""
    tag = buf[i]
    i += 1
    ln = buf[i]
    i += 1
    if ln & 0x80:
        n = ln & 0x7F
        ln = int.from_bytes(buf[i:i + n], "big")
        i += n
    return tag, ln, i


def _der_seq(der):
    """返回 (tag, content_bytes) —— 取第一个 SEQUENCE 的内容。"""
    tag, ln, i = _read_tlv(der, 0)
    return tag, der[i:i + ln]


def to_url(path):
    """把本地路径转成 VSTO 使用的 file:/// URL（与 Office 写入的格式一致）。

    注意：Office 写出的 URL 盘符/路径可能与磁盘上的实际大小写不同
    （例如实际是 E:\\Code\\...，注册表里却是 e:\\code\\...），
    所以查找已有条目时必须做大小写不敏感比较。
    """
    p = os.path.abspath(path).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p
    return "file:///" + p.lstrip("/")


def _url_key(url):
    """用于比较的规范化键：统一小写 + 统一斜杠方向。"""
    return (url or "").replace("\\", "/").rstrip("/").lower()


def find_entry(url):
    """按 Url 查找已有的 Inclusion 条目名（大小写不敏感）。"""
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, INCLUSION_KEY)
    except OSError:
        return None
    try:
        i = 0
        while True:
            try:
                name, val, _t = winreg.EnumValue(k, i)
            except OSError:
                break
            if _url_key(val) == _url_key(url):
                return name
            i += 1
    finally:
        winreg.CloseKey(k)
    return None


def cmd_add(vsto_path):
    if not os.path.isfile(vsto_path):
        raise SystemExit("ERROR: not found: %s" % vsto_path)
    url = to_url(vsto_path)
    pubkey = read_vsto_public_key(vsto_path)

    existing = find_entry(url)
    if existing:
        # 已存在：核对公钥是否与当前清单一致（历史遗留的条目可能写错了公钥），
        # 不一致就覆盖，否则 Office 会因公钥不匹配而拒绝加载。
        try:
            sk = winreg.OpenKey(winreg.HKEY_CURRENT_USER, INCLUSION_KEY + "\\" + existing)
            old = winreg.QueryValueEx(sk, "PublicKey")[0]
            winreg.CloseKey(sk)
        except OSError:
            old = None
        if old == pubkey:
            print("already in inclusion list (public key matches): %s" % existing)
            return 0
        print("updating existing entry with correct public key: %s" % existing)
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, INCLUSION_KEY + "\\" + existing, 0, winreg.KEY_WRITE)
        try:
            winreg.SetValueEx(k, "Url", 0, winreg.REG_SZ, url)
            winreg.SetValueEx(k, "PublicKey", 0, winreg.REG_SZ, pubkey)
        finally:
            winreg.CloseKey(k)
        return 0

    name = "{%s}" % str(uuid.uuid4()).upper()
    k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, INCLUSION_KEY + "\\" + name, 0, winreg.KEY_WRITE)
    try:
        winreg.SetValueEx(k, "Url", 0, winreg.REG_SZ, url)
        winreg.SetValueEx(k, "PublicKey", 0, winreg.REG_SZ, pubkey)
    finally:
        winreg.CloseKey(k)
    print("added to inclusion list:")
    print("  %s" % name)
    print("  Url = %s" % url)
    return 0


def cmd_remove(vsto_path):
    url = to_url(vsto_path)
    name = find_entry(url)
    if not name:
        print("not in inclusion list")
        return 0
    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, INCLUSION_KEY + "\\" + name)
    print("removed: %s" % name)
    return 0


def cmd_list():
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, INCLUSION_KEY)
    except OSError:
        print("(no inclusion list)")
        return 0
    try:
        i = 0
        while True:
            try:
                name, val, _t = winreg.EnumValue(k, i)
            except OSError:
                break
            if val and val.startswith("file:"):
                print("  %s" % val)
            i += 1
    finally:
        winreg.CloseKey(k)
    return 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    cmd = sys.argv[1].lower()
    if cmd == "list":
        return cmd_list()
    if cmd in ("add", "remove"):
        if len(sys.argv) < 3:
            print("usage: vsto_inclusion.py %s <path-to-vsto>" % cmd)
            return 2
        return cmd_add(sys.argv[2]) if cmd == "add" else cmd_remove(sys.argv[2])
    print("unknown command: %s" % cmd)
    return 2


if __name__ == "__main__":
    sys.exit(main())
