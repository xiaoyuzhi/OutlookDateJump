using System;
using Microsoft.Office.Core;

namespace OutlookDateJump
{
    /// <summary>
    /// Outlook 日期跳转加载项入口（用户代码部分）。
    /// 通过 Ribbon XML 在日历功能区“跳转日期”按钮，点击弹出 WPF 日期选择面板后调用 View.GoToDate 跳转。
    /// 运行时管线（构造、Application 绑定、Globals）见 ThisAddIn.Designer.cs。
    /// </summary>
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 预热 WPF（推迟到启动后的空闲时刻），消除首次点击「跳转日期」的冷启动卡顿
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                new Action(WpfHelper.Prewarm),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 备注: Outlook不会再触发这个事件。
        }

        /// <summary>
        /// 返回功能区扩展对象（Ribbon XML 方式）。
        /// </summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
