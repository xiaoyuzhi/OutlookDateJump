# 第三方组件与许可声明（Third-Party Notices）

本项目使用了以下第三方组件，各组件的版权与许可归其作者所有。分发本项目的源码或二进制时，
请一并保留本文件。

---

## HandyControl 3.5.1

- 项目主页：<https://github.com/HandyOrg/HandyControl>
- 许可证：MIT License
- 版权：Copyright (c) 2018-2024 HandyOrg

本项目以**相对路径引用随仓库分发的 HandyControl.dll**（`lib/HandyControl.dll`）作为 WPF 界面组件库
（`DatePickerWindow` 的窗口与下拉框样式）。这样做是为了让任何人克隆仓库后无需 NuGet 还原即可直接编译；
其 MIT 许可证全文同时保存在 `lib/HandyControl-LICENSE.txt`，分发时按 MIT 要求附带以下完整许可证文本：

```text
MIT License

Copyright (c) 2018-2024 HandyOrg

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Microsoft 组件（不随源码仓库分发）

本项目在**编译期引用**、或在**运行时依赖**以下 Microsoft 组件。它们不包含在本仓库的源码中，
由用户环境或安装程序的引导程序（Bootstrapper）提供，受 Microsoft 各自的许可协议约束：

- .NET Framework 4.7.2
- Microsoft Visual Studio 2010 Tools for Office Runtime（VSTO Runtime 4.0）
- Microsoft Office 主互操作程序集（Office PIAs）——以 `EmbedInteropTypes=true` 方式嵌入本程序集，**不单独分发** PIA 文件
- `Microsoft.Office.Tools.Common` / `Microsoft.Office.Tools.Outlook` 运行时程序集

> 说明：这些组件是构建/运行 Outlook VSTO 加载项的必要环境，其版权归 Microsoft Corporation 所有，
> 本项目仅以微软允许的方式引用，不主张对其的任何权利。请在分发二进制安装包时，确保这些组件
> 通过官方引导程序安装，而不是把微软程序集直接打包进源码仓库。

---

## 图标与图片素材

全部界面图标均由作者使用 AI 工具生成，作者拥有其使用权，不涉及第三方素材的授权问题。

| 用途 | 源文件 | 说明 |
|------|--------|------|
| **程序图标**（PowerShell 版窗口图标） | `LOGO.png`（508×508，RGBA，透明底） | `OutlookDateJump.ps1` 按脚本同目录相对路径加载；与 `.ps1` 一起拷贝即可移植 |
| **Outlook 功能区按钮 + 弹窗标题栏图标** | `Calendar.png`（1536×1536，RGBA，透明底） | 经 `tools/make_icon.ps1` 裁剪透明边 → 生成 64×64 的 `OutlookDateJump/CalendarIcon.png`；`Ribbon1.xml` 的 `image="CalendarIcon"` 与 `DatePickerWindow` 的 32×32 图标都使用它 |

> 说明：`CalendarIcon.png` 是唯一被嵌入程序集的图标资源。更换按钮/标题栏图标，
> 只需替换源文件后重跑 `tools\make_icon.ps1`，无需改动代码。
