# Outlook 日期跳转加载项 — 开发规划

> 项目目标：在经典版 Outlook 日历的功能区加一个「跳转日期」按钮，点击弹出**年/月可直接点选的日历面板**，选中日期后日历视图立即跳转，替代内置「转到日期」对话框中只能手动录入日期的交互。

## 1. 背景与可行性结论

- 内置「转到日期」对话框的日期下拉只能浏览**当前年**的月份，换年份只能靠小箭头，慢。
- 经调研，Outlook 对象模型（COM/VSTO 通用）提供 **`View.GoToDate(DateTime)`** 方法，可直接让日历视图跳到任意日期，**无需操作内置对话框**（官方文档：learn.microsoft.com → `View.GoToDate method (Outlook)`）。
- 截图确认为经典版桌面 Outlook（有「文件/发送/接收」菜单），支持 VSTO 加载项。⚠️ 新版 Outlook（"新 Outlook"/olk.exe）不支持 COM/VSTO/VBA，本方案仅覆盖经典版。

## 2. 技术选型

| 项 | 选择 | 说明 |
|---|---|---|
| 宿主 | 经典版 Outlook 2016/2019/365 桌面版 | 不支持新版 Outlook |
| 框架 | VSTO Add-in (C#) | 功能区按钮原生集成 |
| UI | WPF 窗口（在 VSTO 中承载） | 自绘日历面板，控件体验可控 |
| 目标框架 | .NET Framework 4.7.2+ | VSTO 要求 |
| IDE | Visual Studio 2022 Community（勾选"Office/SharePoint 开发"工作负载） | 提供 VSTO 项目模板 |
| 部署 | ClickOnce 发布 | 单机/小范围安装最省事 |

## 3. 方案设计

### 3.1 功能区按钮

- 用 Ribbon Designer 添加「跳转日期」按钮（图标可用内置 Calendar 图标），放在日历资源管理器的 **ExplorerRibbon**。
- 通过 `GetVisible`/`GetEnabled` 回调：仅当当前文件夹为日历（`CurrentFolder.DefaultItemType == olAppointmentItem`）时显示/启用；在邮件等视图自动隐藏，避免误点。

### 3.2 核心跳转逻辑（约 10 行）

```csharp
private void GoToDate(DateTime target)
{
    var explorer = Application.ActiveExplorer();
    if (explorer?.CurrentView is Outlook.CalendarView cv)
    {
        cv.GoToDate(target);   // 日历视图跳到目标日期
    }
}
```

要点：必须取 `Explorer.CurrentView`（而不是 `Folder.CurrentView`），否则不生效。

### 3.3 日历面板 UI（WPF）

弹一个置顶小窗（无任务栏图标、失焦自动关闭可选）：

- **标题行**：`« 2026 ‹` + `7月 › ›` —— 年份用 «/» 或点年份弹 2000–2050 下拉，月份用 ‹/›，**换年只需一次点击/一次滚动**。
- **月网格**：6×7 星期网格，点击任意日期即关闭面板并执行 `GoToDate`。
- **底部快捷键**：`今天`、`明天`、`下周一`（可选，二期）。
- 默认选中当前显示日期；支持滚轮/方向键微调（可选）。

### 3.4 边界与细节

- 若当前视图不是 CalendarView（如"日程视图"），先 `cv = explorer.CurrentView as CalendarView`，取不到时回退：切换到日历视图再跳，或给出提示。
- 目标日期保持用户当前的视图模式（日/周/月），只改日期不改模式——`GoToDate` 本身即如此，不要额外设置 `CalendarViewMode`。
- 面板记住"上次跳转日期"，下次打开默认定位到该月（存 `Settings`）。

## 4. 实施步骤（里程碑）

1. **POC 验证（0.5h）**：先用 PowerShell/Python COM 一行代码验证 `GoToDate` 在本机 Outlook 上生效：
   ```powershell
   $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
   $ol.ActiveExplorer().CurrentView.GoToDate([datetime]"2026-12-25")
   ```
2. **环境与空壳（0.5h）**：装 VS2022 + Office 开发负载；新建 Outlook VSTO 加载项，确认 Outlook 启动时能加载（显示测试提示）。
3. **功能区按钮（1h）**：加「跳转日期」按钮 + 日历文件夹可见性回调。
4. **WPF 日历面板（3-5h）**：年月点选 + 月网格 + 选中跳转。
5. **联调（1h）**：各种视图模式（日/周/月/日程视图）、多日历文件夹、异常处理。
6. **打包发布（1h）**：ClickOnce 发布，设置签名（可选），安装验证。

**预计总工作量：1 个工作日内。**

## 5. 风险与备选

| 风险 | 应对 |
|---|---|
| 机器没装 VSTO 运行时 | ClickOnce 会自动带 VSTO Runtime 引导安装 |
| 公司策略禁用宏/加载项 | VSTO 属于受信任的加载项机制，一般不受 VBA 宏策略限制；如有组策略锁死，回退到「独立小工具+全局热键」方案（不改 Outlook，COM 调用相同） |
| 换了新版 Outlook | 新版不支持任何本地扩展，届时需改用 Graph API / 浏览器方案，超出本项目范围 |
| Outlook 未启动 | 按钮本身在 Outlook 内，不存在此问题 |

## 6. 二期可选增强

- 常用日期快捷键：今天 / 明天 / 本周末 / 下周一 / 月底
- 键盘直达：面板内输入 `2026-12-25` 或 `1225` 快速定位
- 同时支持多个日历账户切换
