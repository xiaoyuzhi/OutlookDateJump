#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打包脚本：产出「安装版」与「绿色版」两个可分发程序包。

用法（仓库根目录下）：
    python tools/make_packages.py

前置条件：已运行 OutlookDateJump\\build_publish.cmd（构建 + ClickOnce 发布），
产物位于 OutlookDateJump\\bin\\Release\\。本脚本只做「组装 + 压缩」，不编译。

产出（dist\\ 目录下）：
    OutlookDateJump-<ver>-Setup.zip     安装版：ClickOnce setup.exe + .vsto + Application Files
    OutlookDateJump-<ver>-Portable.zip  绿色版：|vstolocal 便携目录 + 注册/卸载脚本

两个包都不含任何本机绝对路径：绿色版注册脚本按「脚本自身所在目录」解析
vsto 路径，因此整个目录可以搬到任意机器、任意路径使用。
"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = os.path.join(ROOT, "OutlookDateJump", "bin", "Release")
PUBLISH = os.path.join(BIN, "app.publish")
DIST = os.path.join(ROOT, "dist")

# 绿色版需要随包携带的可执行/依赖文件（清单里以相对 codebase 引用）
PORTABLE_FILES = [
    "OutlookDateJump.dll",
    "OutlookDateJump.vsto",
    "OutlookDateJump.dll.manifest",
    "HandyControl.dll",
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
    "Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll",
]

REGISTER_PS1 = r"""# 安装（绿色版）：把本目录注册为当前用户的 Outlook 加载项。
# 用法：右键本文件 -> 使用 PowerShell 运行；或：
#   powershell -ExecutionPolicy Bypass -File .\Register-ODJ.ps1
# 注册后需重启 Outlook 才会生效。
#
# 注意：除了写 Outlook 的加载项注册表键，还必须把清单加入 VSTO 的
# 「受信任清单」(Inclusion List, HKCU\Software\Microsoft\VSTO\Security\Inclusion)。
# 缺少这一项时，Office 会静默拒绝加载（功能区没有任何按钮、事件 ID 45 也不列出），
# 是最常见的「装上了却不生效」原因。
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vsto = Join-Path $dir 'OutlookDateJump.vsto'
if (-not (Test-Path $vsto)) {
    Write-Host "ERROR: not found: $vsto" -ForegroundColor Red
    Write-Host "Please keep this script in the same folder as OutlookDateJump.vsto." -ForegroundColor Red
    exit 1
}
$key = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookDateJump'
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
# vstolocal：manifest 直接指向本目录的 .vsto，不经 ClickOnce 缓存
$manifest = 'file:///' + ($vsto -replace '\\','/') + '|vstolocal'
Set-ItemProperty -Path $key -Name 'Manifest'  -Value $manifest -Type String
Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Outlook Date Jump' -Type String
Set-ItemProperty -Path $key -Name 'Description'  -Value 'OutlookDateJump' -Type String
Write-Host 'OK. Registered:' -ForegroundColor Green
Write-Host "  $key"
Write-Host "  Manifest = $manifest"

# --- 追加到 VSTO 受信任清单 ---
$inc = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
$url = 'file:///' + ($vsto -replace '\\','/')
$found = $false
if (Test-Path $inc) {
    Get-ChildItem $inc | ForEach-Object {
        $u = (Get-ItemProperty $_.PSPath).Url
        if ($u -eq $url) { $found = $true }
    }
}
if ($found) {
    Write-Host '  Inclusion list: already trusted' -ForegroundColor DarkGray
} else {
    Write-Host '  Adding to VSTO inclusion list (trust) ...' -ForegroundColor Yellow
    $helper = Join-Path $dir 'vsto_inclusion.py'
    if (Test-Path $helper) {
        $py = $null
        foreach ($cand in @('python', 'py')) {
            $cmd = Get-Command $cand -ErrorAction SilentlyContinue
            if ($cmd) { $py = $cmd.Source; break }
        }
        if ($py) {
            & $py $helper add $vsto
            if ($LASTEXITCODE -ne 0) {
                Write-Host '  WARNING: could not add to inclusion list automatically.' -ForegroundColor Yellow
                Write-Host '  If the button does not appear, run tools\vsto_inclusion.py manually,' -ForegroundColor Yellow
                Write-Host '  or import the .cer into Trusted Root Certification Authorities.' -ForegroundColor Yellow
            }
        } else {
            Write-Host '  WARNING: python not found; cannot add to inclusion list automatically.' -ForegroundColor Yellow
        }
    } else {
        Write-Host '  WARNING: vsto_inclusion.py not found next to this script.' -ForegroundColor Yellow
    }
}
Write-Host 'Please restart Outlook to activate.' -ForegroundColor Yellow
"""

UNREGISTER_PS1 = r"""# 卸载（绿色版）：移除当前用户的 Outlook 加载项注册。
# 用法：powershell -ExecutionPolicy Bypass -File .\Unregister-ODJ.ps1
# 注意：只删注册表项，不删除本目录文件；如需彻底清理请手动删除整个文件夹。
$ErrorActionPreference = 'SilentlyContinue'

# 1) 未版本化的加载项键（|vstolocal 部署写入的就是这个）
$key = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookDateJump'
if (Test-Path $key) {
    Remove-Item -Path $key -Recurse -Force
    Write-Host 'OK. Unregistered. Please restart Outlook.' -ForegroundColor Green
} else {
    Write-Host 'Add-in key already absent.'
}

# 2) 版本化的加载项键（曾用 ClickOnce 安装过时会残留；
#    它存在时 Office 会优先走 ClickOnce 部署而忽略本目录的 |vstolocal 注册）
$key16 = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\OutlookDateJump'
if (Test-Path $key16) {
    Remove-Item -Path $key16 -Recurse -Force
    Write-Host 'Removed leftover versioned add-in key.' -ForegroundColor Green
}

# 3) VSTO 受信任清单里的本目录条目
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vsto = Join-Path $dir 'OutlookDateJump.vsto'
$helper = Join-Path $dir 'vsto_inclusion.py'
if ((Test-Path $helper) -and (Test-Path $vsto)) {
    foreach ($cand in @('python', 'py')) {
        $cmd = Get-Command $cand -ErrorAction SilentlyContinue
        if ($cmd) { & $cmd.Source $helper remove $vsto; break }
    }
}

# 4) Office 的「已禁用加载项」记录（加载项异常时会被拉黑，导致按钮消失）
$dis = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems'
if (Test-Path $dis) {
    (Get-ItemProperty $dis).PSObject.Properties |
        Where-Object { $_.Name -notlike 'PS*' } |
        ForEach-Object { Remove-ItemProperty -Path $dis -Name $_.Name -Force }
}
"""

PORTABLE_README = """Outlook Date Jump - 绿色版（免安装 / 便携版）v{VERSION}
================================================================

这是什么
--------
一个「解压即用」的 Outlook 加载项。不需要安装程序，也不需要管理员权限，
不写 ClickOnce 缓存，所有文件都在本文件夹内。

版本
----
v{VERSION}（语义化版本 semver.org；仓库根 VERSION 文件是唯一事实来源）

怎么用
------
1. 把整个文件夹解压到任意位置（例如 D:\\Tools\\OutlookDateJump）。
2. 注册（两种方式，任选其一）：
   a) 右键 Register-ODJ.ps1 -> 「使用 PowerShell 运行」
   b) 在该目录打开 PowerShell，执行：
        powershell -ExecutionPolicy Bypass -File .\\Register-ODJ.ps1
   说明：若直接右键运行提示「禁止运行脚本」，那是系统的执行策略（ExecutionPolicy）
   限制，用上面 b) 的 -ExecutionPolicy Bypass 方式即可，无需修改系统全局设置。
3. 完全退出并重新启动 Outlook（只关窗口不算，必须完全退出）。
4. 切到日历模块的「开始」选项卡，即可看到「跳转」组里的「跳转日期」按钮。

怎么卸载
--------
1. 右键 Unregister-ODJ.ps1 -> 「使用 PowerShell 运行」
   （或 powershell -ExecutionPolicy Bypass -File .\\Unregister-ODJ.ps1）。
2. 重新启动 Outlook。
3. 删除本文件夹即可。

注意事项
--------
* 注册脚本只写当前用户（HKCU），不需要管理员权限，也不会影响其他用户。
* 注册脚本除了写 Outlook 的加载项键，还会把清单加入 VSTO 的「受信任清单」
  (HKCU\\Software\\Microsoft\\VSTO\\Security\\Inclusion)。这一项**必须有**，
  否则 Office 会静默拒绝加载（功能区没有按钮、Outlook 事件 ID 45 也不列出该加载项），
  是「装上了却不生效」最常见的原因。脚本会自动处理；
  若提示找不到 python，可用同目录的 vsto_inclusion.py 手动执行。
* 移动/重命名本文件夹后，请重新运行 Register-ODJ.ps1（脚本按自身所在目录
  解析路径，不含任何写死的绝对路径）。
* 若 Outlook 提示「无法加载加载项」或按钮不出现：
  - 确认已完全重启 Outlook（不是只关窗口）；
  - 确认 Register-ODJ.ps1 运行成功；
  - 检查 HKCU\\Software\\Microsoft\\Office\\Outlook\\Addins\\OutlookDateJump
    的 Manifest 值是否指向本目录的 OutlookDateJump.vsto。
* 加载项被 Office 拉黑时（按钮整组消失），把该键下的 LoadBehavior 改为 3，
  并删除 HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Outlook\\Resiliency
  下的 DisabledItems 项，然后重启 Outlook。

文件说明
--------
  OutlookDateJump.dll                 加载项主程序
  OutlookDateJump.vsto                清单（注册脚本指向的就是它）
  OutlookDateJump.dll.manifest        依赖清单（相对路径引用，故目录可整体搬移）
  HandyControl.dll                    WPF UI 组件库（MIT 许可）
  Microsoft.Office.Tools.*.Utilities.dll  VSTO 运行时辅助库
  Register-ODJ.ps1                    注册脚本
  Unregister-ODJ.ps1                  卸载脚本
"""

SETUP_README = """Outlook Date Jump - 安装版（ClickOnce）v{VERSION}
=========================================

这是什么
--------
标准安装包，双击 setup.exe 即可安装。可从「控制面板 -> 程序和功能」卸载。

版本
----
v{VERSION}（语义化版本 semver.org；仓库根 VERSION 文件是唯一事实来源）

怎么安装
--------
1. 把本文件夹解压到任意位置（路径不要包含特殊字符）。
2. 双击 setup.exe，按提示完成安装。
3. 启动 Outlook（首次安装后建议重启一次 Outlook 让加载项被激活）。
4. 切到日历模块的「开始」选项卡，即可看到「跳转」组里的「跳转日期」按钮。

重要：本包使用自签名证书
------------------------
签名证书 CN=OutlookDateJump 是开发者自签名的，不是受信任的商业证书。
首次安装时 Windows 可能弹出「未知发布者」提示：
  - 点击「更多信息」->「仍要运行」即可继续；
  - 若要消除提示，可双击同目录的 OutlookDateJump.cer，
    按向导将其导入「受信任的根证书颁发机构」（需管理员权限）。
    导入后请重启 setup.exe 再安装。

运行前提（缺一不可）
--------------------
* Microsoft Outlook 2016 / 2019 / 2021 / 365（桌面版）。
* Microsoft .NET Framework 4.7.2 或更高版本。
* Microsoft Visual Studio 2010 Tools for Office Runtime（VSTO Runtime 4.0）。
  若机器上没有，安装后 Outlook 不会加载本加载项，也不会报错——表现是
  「安装成功但功能区没有按钮」。缺失时请安装 VSTO Runtime Redistributable
  （微软官网提供 x86/x64 两个包，64 位 Outlook 需装 x64 版）。

怎么卸载
--------
控制面板 -> 程序和功能 -> 「Outlook 日期跳转」-> 卸载。

说明
----
* 本包不联网、不自动更新（UpdateEnabled=false），安装来源为本地磁盘。
* 「绿色版」（免安装包）无需任何安装步骤，如需可向作者索取。
"""


def read_version():
    """读取语义化版本（semver.org）：仓库根 VERSION 文件是唯一事实来源。

    格式 MAJOR.MINOR.PATCH[-预发布标识]，例如 1.4.0 / 1.4.0-rc.1。
    与 csproj 的解析逻辑保持一致（csproj 用 MSBuild 的 ReadAllText+Split 拆解）。
    """
    p = os.path.join(ROOT, "VERSION")
    with open(p, "r", encoding="utf-8-sig") as f:
        raw = f.read().strip()
    m = re.match(r"^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?$", raw)
    if not m:
        sys.exit("ERROR: VERSION file is not a valid SemVer string: %r" % raw)
    return raw


def read_build_revision():
    """ClickOnce 部署修订号（tools/revision.txt，单调递增的构建计数）。

    注意：这不是「版本号」，而是 ClickOnce 用来判断是否需要重新部署的计数器。
    版本号由 VERSION 文件（SemVer）提供，两者职责不同，不要混淆。
    """
    p = os.path.join(ROOT, "tools", "revision.txt")
    with open(p, "r", encoding="ascii") as f:
        return f.read().strip()


def find_cert_cer():
    """在 dist 旁边找导出的 .cer（可选）。找不到就跳过，不影响打包。"""
    for name in ("OutlookDateJump.cer",):
        for base in (ROOT, DIST):
            p = os.path.join(base, name)
            if os.path.isfile(p):
                return p
    return None


def zip_dir(src_dir, zip_path):
    """把目录压成 zip（相对路径为根，便于解压后直接用）。"""
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for base, _dirs, files in os.walk(src_dir):
            for fn in files:
                full = os.path.join(base, fn)
                rel = os.path.relpath(full, src_dir)
                zf.write(full, rel.replace("\\", "/"))
    return zip_path


def write_text(path, text, bom=False):
    """写文本文件。中文说明用 UTF-8 BOM，Windows 记事本才不会乱码。"""
    enc = "utf-8-sig" if bom else "utf-8"
    with open(path, "w", encoding=enc, newline="\r\n") as f:
        f.write(text)


def build_portable(version):
    """绿色版：整目录可搬移 + 注册/卸载脚本。

    注意：解压后的**目录名不含版本号**（固定为 OutlookDateJump-Portable），
    这样每次升版后用户重新注册时，VSTO 的 Security\\Inclusion 白名单里
    仍是同一个路径，不会因为改名而被 Office 静默拒绝加载。
    版本号体现在 zip 文件名与包内 VERSION 文件里。
    """
    name = "OutlookDateJump-Portable"
    out = os.path.join(DIST, name)
    if os.path.isdir(out):
        shutil.rmtree(out)
    os.makedirs(out)

    missing = [f for f in PORTABLE_FILES if not os.path.isfile(os.path.join(BIN, f))]
    if missing:
        sys.exit("ERROR: missing build outputs: %s\nRun build_publish.cmd first." % missing)

    for f in PORTABLE_FILES:
        shutil.copy2(os.path.join(BIN, f), os.path.join(out, f))

    # VSTO 受信任清单工具：注册脚本用它把 .vsto 加入 Inclusion List
    inc_helper = os.path.join(ROOT, "tools", "vsto_inclusion.py")
    if os.path.isfile(inc_helper):
        shutil.copy2(inc_helper, os.path.join(out, "vsto_inclusion.py"))

    write_text(os.path.join(out, "Register-ODJ.ps1"), REGISTER_PS1)
    write_text(os.path.join(out, "Unregister-ODJ.ps1"), UNREGISTER_PS1)
    write_text(os.path.join(out, "使用说明.txt"), PORTABLE_README.replace("{VERSION}", version), bom=True)
    write_text(os.path.join(out, "VERSION"), version + "\n", bom=False)

    cer = find_cert_cer()
    if cer:
        shutil.copy2(cer, os.path.join(out, "OutlookDateJump.cer"))

    z = zip_dir(out, os.path.join(DIST, "OutlookDateJump-v%s-Portable.zip" % version))
    print("portable -> %s" % z)
    return z, out


def build_setup(version):
    """安装版：ClickOnce 发布目录整体打包。"""
    name = "OutlookDateJump-v%s-Setup" % version
    out = os.path.join(DIST, name)
    if os.path.isdir(out):
        shutil.rmtree(out)
    if not os.path.isdir(PUBLISH):
        sys.exit("ERROR: %s not found. Run build_publish.cmd first." % PUBLISH)

    shutil.copytree(PUBLISH, out)

    write_text(os.path.join(out, "使用说明.txt"), SETUP_README.replace("{VERSION}", version), bom=True)
    write_text(os.path.join(out, "VERSION"), version + "\n", bom=False)

    cer = find_cert_cer()
    if cer:
        shutil.copy2(cer, os.path.join(out, "OutlookDateJump.cer"))

    z = zip_dir(out, os.path.join(DIST, "%s.zip" % name))
    print("setup    -> %s" % z)
    return z, out


def main():
    if not os.path.isdir(BIN):
        sys.exit("ERROR: %s not found. Run build_publish.cmd first." % BIN)
    os.makedirs(DIST, exist_ok=True)
    version = read_version()
    print("version = %s" % version)
    zp, _ = build_portable(version)
    zs, _ = build_setup(version)
    for z in (zp, zs):
        print("  %-52s %8.1f KB" % (os.path.basename(z), os.path.getsize(z) / 1024.0))
    print("DONE")


if __name__ == "__main__":
    main()
