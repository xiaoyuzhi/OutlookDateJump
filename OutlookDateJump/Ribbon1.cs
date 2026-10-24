using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookDateJump
{
    /// <summary>
    /// 功能区扩展（Ribbon XML 方式），在日历选项卡内放置“跳转日期”按钮，
    /// 并仅当当前文件夹为日历时启用。
    /// </summary>
    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("OutlookDateJump.Ribbon1.xml");
        }

        #endregion

        #region 回调

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// 仅当当前文件夹为日历时，按钮可用。
        /// </summary>
        public bool GetEnabled(IRibbonControl control)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.CurrentFolder == null)
                {
                    return false;
                }

                return explorer.CurrentFolder.DefaultItemType == Outlook.OlItemType.olAppointmentItem;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// loadImage 回调：Office 把按钮 image="CalendarIcon" 属性的字符串传进来，返回对应图片。
        /// </summary>
        public RibbonImageHelper.IPictureDisp GetRibbonImage(string imageName)
        {
            return RibbonImageHelper.GetIcon(imageName);
        }

        private static void Trace(string msg)
        {
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "odjdbg.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + "\r\n");
            }
            catch { }
        }

        /// <summary>
        /// 点击“跳转日期”：弹出日历面板，确认后跳转到选中日期。
        /// </summary>
        public void OnDateJump(IRibbonControl control)
        {
            Trace("OnDateJump enter");
            var app = Globals.ThisAddIn.Application;
            var explorer = app.ActiveExplorer();
            if (explorer == null)
            {
                Trace("no active explorer");
                return;
            }

            DateTime? picked = null;

            // 把 Outlook 主窗口句柄传给 WPF 弹窗，设为 Owner：
            // 弹窗成为 Outlook 的"拥有窗口"，Z 序正确、关闭后焦点回到 Outlook，
            // 也避免首次真实鼠标点击被系统用于"激活窗口"而被吃掉（表现为要点的两次）。
            // 本加载项运行在 Outlook 进程内，故直接用当前进程主窗口句柄。
            IntPtr owner = IntPtr.Zero;
            try
            {
                using (var p = System.Diagnostics.Process.GetCurrentProcess())
                {
                    owner = p.MainWindowHandle;
                }
            }
            catch { }

            // 弹窗内部只做"选日期 + 关窗"，不执行任何 Outlook COM 调用。
            WpfHelper.ShowCalendarPicker(DateTime.Today, d => picked = d, owner);
            Trace("dialog closed, picked=" + (picked.HasValue ? picked.Value.ToString("yyyy-MM-dd") : "NULL"));

            if (!picked.HasValue)
            {
                return;
            }

            // 此刻模态窗口已完全关闭，Outlook 回到正常消息循环，
            // 立刻跳一次。第一拍成功就到此为止——
            // 重复调用 GoToDate 会让 Outlook 反复重绘日历，表现为画面闪几下。
            bool jumped;
            try
            {
                jumped = JumpToDate(picked.Value, true);
            }
            catch
            {
                jumped = false;
            }

            // 只有第一拍失败时才用 Timer 兜底重试
            // （覆盖"模态收尾期 GoToDate 被丢弃"等历史环境问题）。
            if (!jumped)
            {
                ScheduleJump(picked.Value);
            }
        }

        /// <summary>
        /// 用 WinForms Timer 兜底重试跳转（仅在第一拍失败时启用）。
        /// WinForms Timer 依赖 Win32 消息循环（Outlook UI 线程有），
        /// 即使 WPF Dispatcher 已停止泵消息也会按时触发——这是它相对
        /// WPF DispatcherTimer 的关键优势（后者在模态关闭后根本不触发）。
        /// 注意：第一拍成功时不会走到这里，正常跳转只调用一次 GoToDate，不闪屏。
        /// </summary>
        private static void ScheduleJump(DateTime date)
        {
            try
            {
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 120;
                int attempts = 0;

                timer.Tick += (s, e) =>
                {
                    attempts++;
                    // 重试的第一拍带 Apply()（对齐正常路径），之后只 GoToDate
                    bool ok = JumpToDate(date, attempts == 1);
                    Trace("retry tick #" + attempts + " ok=" + ok);

                    // 重试成功后再补两拍覆盖"视图重绘滞后"；
                    // 此路径本来就是异常兜底，多一两次重绘可接受。
                    if ((ok && attempts >= 3) || attempts >= 10)
                    {
                        timer.Stop();
                        timer.Dispose();
                    }
                };

                timer.Start();
            }
            catch (Exception ex)
            {
                Trace("timer error: " + ex.Message);
            }
        }

        #endregion

        #region 核心逻辑

        /// <summary>
        /// 调用 Outlook 视图的 GoToDate，使日历跳转到目标日期。
        /// GoToDate 定义于 View 基类，因此日程安排视图等非 CalendarView 也能跳转。
        /// 返回 true 表示本次调用未抛异常。
        /// </summary>
        /// <param name="reapply">
        /// 是否先 Apply() 一次视图。只在第一拍做：重复 Apply() 会把视图恢复成
        /// 已保存的定义，反而可能把日历位置拨回去。
        /// </param>
        private static bool JumpToDate(DateTime date, bool reapply)
        {
            try
            {
                // 每次都用「当前活动 Explorer」：弹窗关闭后焦点已回到 Outlook，
                // 活动 Explorer 就是用户正在看的日历窗口，无需 Activate()。
                // （模态期间 Activate() 必然失败，且会抢走弹窗焦点，是"点两次"的帮凶）
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null)
                {
                    Trace("JumpToDate: no active explorer");
                    return false;
                }

                var view = explorer.CurrentView;
                if (view == null)
                {
                    Trace("JumpToDate: CurrentView null");
                    return false;
                }

                var calendarView = view as Outlook.CalendarView;
                if (calendarView != null)
                {
                    // 顺序要点：先 Apply() 让视图就绪，再 GoToDate 定位。
                    // 反过来（先 GoToDate 后 Apply）会把视图重置回原位置。
                    if (reapply)
                    {
                        try { calendarView.Apply(); }
                        catch { }
                    }

                    calendarView.GoToDate(date);
                    Trace("applied+goto " + date.ToString("yyyy-MM-dd") + " reapply=" + reapply);
                }
                else
                {
                    // 其他视图经 COM 晚绑定尝试，不支持会抛异常并被外层捕获
                    Trace("view=" + view.GetType().Name + " (dynamic)");
                    ((dynamic)view).GoToDate(date);
                }

                // 绝不 Marshal.ReleaseComObject(view)：
                // Outlook 会跨调用复用 CurrentView 的 RCW，提前释放后再次取会拿到
                // 已失效的 COM 对象 —— 这正是"第一次跳转无效、第二次才生效"的常见成因。
                return true;
            }
            catch (Exception ex)
            {
                Trace("JumpToDate EXC: " + ex.Message);
                return false;
            }
        }

        #endregion

        #region 资源读取

        private static string GetResourceText(string resourceName)
        {
            var asm = typeof(Ribbon1).Assembly;
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("找不到嵌入资源: " + resourceName);
                }

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        #endregion
    }
}
