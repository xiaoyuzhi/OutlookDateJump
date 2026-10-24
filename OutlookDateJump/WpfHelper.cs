using System;
using System.Windows;

namespace OutlookDateJump
{
    /// <summary>
    /// 在 VSTO（无宿主 WPF Application）中安全地弹出 WPF 窗口。
    /// WPF 的 Application 在每个 AppDomain 只能创建一次：首次创建后静态保留复用，
    /// 绝不调用 Shutdown（否则后续 new Application() 会抛异常导致弹窗只能弹一次）。
    /// </summary>
    internal static class WpfHelper
    {
        private static Application _sharedApp;

        /// <summary>
        /// Outlook 启动时预热 WPF（Application 实例 + 窗口模板 JIT），
        /// 避免首次点击按钮时的冷启动卡顿。
        /// </summary>
        public static void Prewarm()
        {
            try
            {
                EnsureApplication();
                var warm = new DatePickerWindow(DateTime.Today);
                warm.Close();
            }
            catch
            {
                // 预热失败不影响后续正常弹窗
            }
        }

        private static void EnsureApplication()
        {
            if (Application.Current == null && _sharedApp == null)
            {
                _sharedApp = new Application
                {
                    // 默认 OnLastWindowClose 会在弹窗关闭时把 Outlook 主线程的
                    // WPF Dispatcher 一并关掉（Dispatcher 无法重启，导致之后弹窗静默失败）。
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };

                LoadHandyControlTheme();
            }
        }

        /// <summary>
        /// 加载 HandControl 主题资源（皮肤 + 控件样式）。
        /// 资源以 pack URI 形式嵌入 HandyControl.dll 内。
        /// </summary>
        private static void LoadHandyControlTheme()
        {
            try
            {
                var dicts = _sharedApp.Resources.MergedDictionaries;

                // 皮肤（颜色变量）
                dicts.Add(new ResourceDictionary
                {
                    Source = new Uri("pack://application:,,,/HandyControl;component/Themes/SkinDefault.xaml", UriKind.Absolute)
                });

                // 控件样式（Button / ComboBox / Calendar 等）
                dicts.Add(new ResourceDictionary
                {
                    Source = new Uri("pack://application:,,,/HandyControl;component/Themes/Theme.xaml", UriKind.Absolute)
                });
            }
            catch
            {
                // 主题加载失败时回退到系统默认样式，不影响功能
            }
        }

        public static void ShowCalendarPicker(DateTime initial, Action<DateTime> onPicked, IntPtr ownerHwnd)
        {
            EnsureApplication();

            var win = new DatePickerWindow(initial);

            // 设为 Outlook 的拥有窗口：正确的 Z 序 + 激活关系，避免首次点击被系统用于激活。
            if (ownerHwnd != IntPtr.Zero)
            {
                try
                {
                    var helper = new System.Windows.Interop.WindowInteropHelper(win);
                    helper.Owner = ownerHwnd;
                }
                catch
                {
                    // Owner 设置失败不影响功能（回退到无 Owner 的模态窗口）
                }
            }

            // 弹窗内只负责"选日期 + 关窗"，不执行任何 Outlook COM 调用；
            // 跳转由调用方在 ShowDialog 返回后用 WinForms Timer 调度，
            // 从而彻底避开模态消息循环（模态期间 GoToDate 会被静默丢弃）。
            bool? result = win.ShowDialog();

            if (result == true && win.SelectedDate.HasValue)
            {
                onPicked(win.SelectedDate.Value);
            }
        }
    }
}
