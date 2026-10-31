using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace OutlookDateJump
{
    /// <summary>
    /// 年/月/日三个下拉框 + 日历的日期选择面板（HandControl 美化风格）。
    /// 下拉可分别快速选择，且彼此联动（天数随年/月变化，含闰年）。
    /// 选中日期后点“跳转”返回结果。
    /// </summary>
    public partial class DatePickerWindow : HandyControl.Controls.Window
    {
        public DateTime? SelectedDate { get; private set; }

        /// <summary>
        /// 当前版本号（语义化版本 semver.org）。
        /// 取自程序集的 AssemblyInformationalVersion —— 该属性在构建时依据仓库根
        /// VERSION 文件生成（见 csproj 的 GenerateVersionAssemblyInfo 目标），
        /// 因此界面显示的版本与安装包、Git tag 始终一致。
        /// </summary>
        public static string AppVersion
        {
            get
            {
                try
                {
                    var asm = typeof(DatePickerWindow).Assembly;
                    var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
                    if (info != null && !string.IsNullOrEmpty(info.InformationalVersion))
                    {
                        return info.InformationalVersion;
                    }
                    return asm.GetName().Version.ToString();
                }
                catch
                {
                    return "unknown";
                }
            }
        }

        // 防止下拉框与日历互相触发事件造成死循环
        private bool _busy;

        // 防止 MouseDown 与 Click 重复确认
        private bool _confirmed;

        public DatePickerWindow() : this(DateTime.Today)
        {
        }

        public DatePickerWindow(DateTime initial)
        {
            InitializeComponent();

            // 版本号由代码赋值（不用 x:Static：那会在 XAML 解析期强制初始化
            // code-behind 类型，与 InitializeComponent 形成循环，可能导致加载项崩溃）。
            try
            {
                TxtVersion.Text = "v" + AppVersion;
                TxtVersionTip.Text = "版本 v" + AppVersion;
            }
            catch
            {
                // 版本显示失败不影响弹窗其它功能
            }

            LoadTitleIcon();
            InitCombos(initial);

            // 下拉框打开时，把当前选中值尽量滚到可视区中间（而非顶/底边缘）
            CmbYear.DropDownOpened += (s, e) => CenterSelectedInDropDown(CmbYear);
            CmbMonth.DropDownOpened += (s, e) => CenterSelectedInDropDown(CmbMonth);
            CmbDay.DropDownOpened += (s, e) => CenterSelectedInDropDown(CmbDay);

            // 下拉开着时接管鼠标滚轮：弹层能滚（小屏列表超高）就滚内容；
            // 一屏放得下（滑动窗口 11 项全部可见、无可滚内容）就把滚轮转成
            // 「切换选中值」，向上=上一个、向下=下一个，可连续滚到任意年/月/日。
            AddHandler(UIElement.PreviewMouseWheelEvent,
                new System.Windows.Input.MouseWheelEventHandler(Win_PreviewMouseWheel), true);

            // 下拉弹层提速：HandyControl 的 ComboBox 模板里 Popup 自带 Fade 渐显动画
            // + DropShadowEffect 投影，在 Outlook 进程中叠加后弹层出现明显迟滞。
            // 模板加载完成后立即移除这两者，让下拉即时出现。
            CmbYear.Loaded += (s, e) => SpeedUpDropDown(CmbYear);
            CmbMonth.Loaded += (s, e) => SpeedUpDropDown(CmbMonth);
            CmbDay.Loaded += (s, e) => SpeedUpDropDown(CmbDay);

            // 窗口级鼠标兜底（handledEventsToo）：即使鼠标被日历控件捕获、
            // 事件被别的控件标记 Handled，窗口这里依然能看到每一次按下/抬起，
            // 从而可以按物理坐标判断指针是否压在「跳转」上，保证一次点击即生效。
            AddHandler(UIElement.PreviewMouseDownEvent, new System.Windows.Input.MouseButtonEventHandler(Win_PreviewMouseDown), true);
            AddHandler(UIElement.PreviewMouseUpEvent, new System.Windows.Input.MouseButtonEventHandler(Win_PreviewMouseUp), true);

            // 弹窗显示后强制成为前台窗口：
            // 否则用户对弹窗的「第一次」鼠标点击会被 Windows 用于激活窗口而丢失，
            // 表现为"点一次没反应、点第二次才生效"。
            this.Loaded += (s, e) =>
            {
                try
                {
                    this.Activate();
                    this.Focus();
                    System.Windows.Input.Keyboard.Focus(this);
                }
                catch { }

                // 等布局完成后再抢前台，成功率高得多
                Dispatcher.BeginInvoke(
                    new Action(() =>
                    {
                        ForceForeground(new System.Windows.Interop.WindowInteropHelper(this).Handle);
                    }),
                    DispatcherPriority.Loaded);
            };

            // 首帧渲染完成后才拿到真实的布局尺寸，此时再按日历比例定窗口大小，
            // 保证日历铺满、四周不留大片空白。
            this.ContentRendered += (s, e) => FitWindowToCalendar();

            // 窗口被拖动改变大小时，让日历框始终等于日历区域的可用尺寸
            CalBorder.SizeChanged += (s, e) => SyncCalendarBox();
        }

        /// <summary>
        /// 窗口句柄创建后立即抢前台。
        /// </summary>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ForceForeground(new System.Windows.Interop.WindowInteropHelper(this).Handle);
        }

        #region 布局：日历铺满窗体

        /// <summary>日历控件的显示字号（缩放后实际渲染出来的字号）。</summary>
        private const double CalFontSize = 15.0;

        /// <summary>
        /// 日历整体缩放比。WPF 的日历格子尺寸是模板里写死的（约 37px），
        /// 改 FontSize 并不会让它变小，所以用 Viewbox 整体等比缩放，
        /// 同时把控件字号预放大到 CalFontSize/CalZoom，保证显示出来正好是目标字号。
        /// 1.0 = 按日历自然尺寸原样显示（不再缩放）。
        /// </summary>
        private const double CalZoom = 1.0;

        /// <summary>
        /// 在日历自然宽度之外额外增加的宽度：窗口最终比日历「自然宽度」宽出这许多像素，
        /// 且这加宽量由日历整体等比缩放吃掉 —— 日历铺满整个日历区域、不留两侧留白；
        /// 年/月/日下拉与底部按钮本来就随窗口宽度拉伸，三者一起占满空间。
        /// </summary>
        private const double WindowExtraWidth = 70.0;

        /// <summary>
        /// FitWindowToCalendar 实际算出的缩放比（填满加宽窗口时 &gt; 1，即放大）。
        /// SyncCalendarBox 以它为上限：首帧布局时它还是 CalZoom(=1.0)，
        /// 避免日历在正式定尺寸前被瞬间放到很大再缩回（闪烁）。
        /// </summary>
        private double _calScale = CalZoom;

        /// <summary>日历外框的边框 + 内边距（横向合计）。</summary>
        private double BorderPadH
        {
            get
            {
                return CalBorder.BorderThickness.Left + CalBorder.BorderThickness.Right
                     + CalBorder.Padding.Left + CalBorder.Padding.Right;
            }
        }

        /// <summary>日历外框的边框 + 内边距（纵向合计）。</summary>
        private double BorderPadV
        {
            get
            {
                return CalBorder.BorderThickness.Top + CalBorder.BorderThickness.Bottom
                     + CalBorder.Padding.Top + CalBorder.Padding.Bottom;
            }
        }

        /// <summary>
        /// 把 Viewbox 的尺寸设成日历「自然尺寸 × _calScale」与可用空间的较小者。
        /// WPF 的 Viewbox 在本窗体里不会自动横向拉伸（实测会缩成内容的自然宽度，
        /// 两侧各留几十像素空白），所以这里直接给它定死宽高。
        /// 放大上限取 FitWindowToCalendar 算出的 _calScale：日历铺满加宽后的窗口，
        /// 控件字号已由 Cal.FontSize = CalFontSize / scale 预缩放，显示出来仍是目标字号。
        /// </summary>
        private void SyncCalendarBox()
        {
            try
            {
                double w = CalBorder.ActualWidth - BorderPadH;
                double h = CalBorder.ActualHeight - BorderPadV;
                if (w <= 1 || h <= 1)
                {
                    return;
                }

                // 日历在 Viewbox 中按无限约束测量，DesiredSize 即它的“自然尺寸”
                double natW = Cal.DesiredSize.Width;
                double natH = Cal.DesiredSize.Height;
                if (natW > 1 && natH > 1)
                {
                    double s = Math.Min(_calScale, Math.Min(w / natW, h / natH));
                    w = Math.Round(natW * s);
                    h = Math.Round(natH * s);
                }

                CalBox.Width = w;
                CalBox.Height = h;
            }
            catch
            {
                // 布局过程中的瞬时异常可忽略
            }
        }

        /// <summary>
        /// 按日历自身的长宽比反推窗口尺寸，让「日历区域」与日历完全同比例，
        /// 从而铺满且不出现上下或左右的大片空白。
        /// 窗口在日历自然宽度之外加宽 WindowExtraWidth，该加宽量由日历等比缩放吃掉，
        /// 下拉框与按钮随宽度拉伸，整个窗体内部不留空白。
        /// </summary>
        private void FitWindowToCalendar()
        {
            try
            {
                // Viewbox 用无限约束测量子元素，DesiredSize 就是日历的自然尺寸
                double natW = Math.Max(1.0, Cal.DesiredSize.Width);
                double natH = Math.Max(1.0, Cal.DesiredSize.Height);

                // 除日历区域外的固定占用（标题栏、边距、下拉行、按钮行、边框）
                double chromeW = this.ActualWidth - (CalBorder.ActualWidth - BorderPadH);
                double chromeH = this.ActualHeight - (CalBorder.ActualHeight - BorderPadV);
                if (chromeW < 0) { chromeW = 0; }
                if (chromeH < 0) { chromeH = 0; }

                var wa = SystemParameters.WorkArea;

                // 可用区域按「屏幕」算，而不是当前窗口宽度 ——
                // 否则 XAML 里的初始 Width 会把日历限制住，缩放算出 0.99 之类的零碎值。
                double availW = Math.Max(240, wa.Width * 0.95 - chromeW);
                double availH = Math.Max(240, wa.Height * 0.92 - chromeH);

                // 缩放系数：窗口要在日历自然宽度之外再加宽 WindowExtraWidth，
                // 这加宽量由日历等比放大吃掉（日历铺满、下拉/按钮跟随拉伸一起占满），
                // 因此基准缩放 = (natW + WindowExtraWidth) / natW（≥1，即放大）。
                double scale = Math.Max(CalZoom, (natW + WindowExtraWidth) / natW);
                if (natW * scale > availW) { scale = Math.Min(scale, availW / natW); }
                if (natH * scale > availH) { scale = Math.Min(scale, availH / natH); }
                if (scale < 0.6) { scale = 0.6; }

                // 记住本次缩放，供 SyncCalendarBox（SizeChanged 兜底）使用
                _calScale = scale;

                // Viewbox 会把日历连同文字一起缩放 scale 倍，
                // 因此控件自身字号要预先放大到 CalFontSize/scale，显示出来才是目标字号。
                Cal.FontSize = CalFontSize / scale;

                double boxW = Math.Round(natW * scale);
                double boxH = Math.Round(natH * scale);

                // 加宽量已经含在缩放里（boxW = natW × scale ≥ natW + WindowExtraWidth），
                // 这里不再另加 WindowExtraWidth，否则窗口会比预期再宽出一截。
                double newW = boxW + chromeW;
                double newH = boxH + chromeH;

                double dw = newW - this.ActualWidth;
                double dh = newH - this.ActualHeight;

                this.Width = newW;
                this.Height = newH;

                // 以原来中心为锚点缩放，避免窗口跳到屏幕外
                this.Left -= dw / 2.0;
                this.Top -= dh / 2.0;
                if (this.Left < wa.Left) { this.Left = wa.Left; }
                if (this.Top < wa.Top) { this.Top = wa.Top; }
                if (this.Left + newW > wa.Right) { this.Left = wa.Right - newW; }
                if (this.Top + newH > wa.Bottom) { this.Top = wa.Bottom - newH; }

                // 立刻把 Viewbox 尺寸对齐（不依赖后续的 SizeChanged 时机）
                CalBox.Width = boxW;
                CalBox.Height = boxH;

                Trace(string.Format(
                    "fit: nat={0:F0}x{1:F0} box={2:F0}x{3:F0} chrome={4:F0}x{5:F0} scale={6:F2} window={7:F0}x{8:F0} work={9:F0}x{10:F0}",
                    natW, natH, boxW, boxH, chromeW, chromeH, scale, newW, newH, wa.Width, wa.Height));
            }
            catch (Exception ex)
            {
                Trace("fit error: " + ex.Message);
            }
        }

        #endregion

        #region 强制前置

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr pid);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)] bool fAttach);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        /// <summary>
        /// 把自己的窗口强制提到前台并取得键盘焦点。
        /// 单纯 SetForegroundWindow 在别的进程持有前台时会失败，
        /// 因此先用 AttachThreadInput 把本线程输入队列接到当前前台线程，再设置。
        /// </summary>
        private static void ForceForeground(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            try
            {
                // 瞬时置顶再取消：把窗口钉到 Z 序最上层，但不残留 Topmost 属性
                SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

                if (GetForegroundWindow() == hwnd)
                {
                    SetFocus(hwnd);
                    return;
                }

                uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
                uint myThread = GetCurrentThreadId();
                bool attached = false;

                if (fgThread != 0 && fgThread != myThread)
                {
                    attached = AttachThreadInput(myThread, fgThread, true);
                }

                try
                {
                    SetForegroundWindow(hwnd);
                    SetFocus(hwnd);
                }
                finally
                {
                    if (attached)
                    {
                        AttachThreadInput(myThread, fgThread, false);
                    }
                }
            }
            catch
            {
                // 抢前台失败不致命：最坏情况是首次点击用于激活窗口
            }
        }

        #endregion

        #region 窗口级鼠标兜底

        /// <summary>
        /// 窗口隧道阶段的鼠标按下：只记录，不做确认。
        /// 确认放在“抬起”时做——若在按下时就关窗，鼠标抬起会落到下面的 Outlook 窗口上，
        /// Outlook 收到这次点击可能把日历滚到别处，看起来就像“跳了又没跳”。
        /// </summary>
        private void Win_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            bool overOk = IsOverElement(e, BtnOk);
            Trace("winMouseDown src=" + SrcName(e) + " overOk=" + overOk + " captured=" + CaptureName());
        }

        /// <summary>
        /// 窗口隧道阶段的鼠标抬起：指针在「跳转」按钮上就确认。
        /// 用 handledEventsToo:true 注册，因此即使事件被别的控件吞掉/鼠标被日历捕获，
        /// 这里依然会被调用 —— 保证「一次点击」必定生效。
        /// </summary>
        private void Win_PreviewMouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_confirmed)
            {
                return;
            }

            bool overOk = IsOverElement(e, BtnOk);
            Trace("winMouseUp src=" + SrcName(e) + " overOk=" + overOk + " captured=" + CaptureName());
            if (overOk)
            {
                ConfirmJump();
            }
        }

        /// <summary>
        /// 判断本次鼠标事件的物理位置是否落在指定元素上。
        /// 用命中测试而不是 OriginalSource，因为鼠标被捕获时
        /// OriginalSource 会是被捕获的元素，而不是指针真正压着的元素。
        /// </summary>
        private bool IsOverElement(System.Windows.Input.MouseEventArgs e, DependencyObject target)
        {
            try
            {
                var pos = e.GetPosition(this);
                if (double.IsNaN(pos.X) || double.IsNaN(pos.Y))
                {
                    return false;
                }

                var hit = this.InputHitTest(pos) as DependencyObject;
                return IsSelfOrDescendant(hit, target);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsSelfOrDescendant(DependencyObject node, DependencyObject ancestor)
        {
            while (node != null)
            {
                if (ReferenceEquals(node, ancestor))
                {
                    return true;
                }

                DependencyObject parent;
                if (node is Visual || node is System.Windows.Media.Media3D.Visual3D)
                {
                    parent = VisualTreeHelper.GetParent(node);
                }
                else
                {
                    parent = LogicalTreeHelper.GetParent(node);
                }

                node = parent;
            }

            return false;
        }

        /// <summary>
        /// 判断命中元素（或其某个祖先）是否指定类型，用于区分"双击落在日期格子"
        /// 还是落在表头/箭头/空白（那些不是 CalendarDayButton）。
        /// </summary>
        private static bool IsSelfOrDescendantOfType(DependencyObject node, Type type)
        {
            while (node != null)
            {
                if (type.IsInstanceOfType(node))
                {
                    return true;
                }

                DependencyObject parent;
                if (node is Visual || node is System.Windows.Media.Media3D.Visual3D)
                {
                    parent = VisualTreeHelper.GetParent(node);
                }
                else
                {
                    parent = LogicalTreeHelper.GetParent(node);
                }

                node = parent;
            }

            return false;
        }

        private static string CaptureName()
        {
            try
            {
                var cap = System.Windows.Input.Mouse.Captured;
                return cap == null ? "null" : cap.GetType().Name;
            }
            catch
            {
                return "?";
            }
        }

        private static string SrcName(System.Windows.RoutedEventArgs e)
        {
            var fe = e.OriginalSource as FrameworkElement;
            if (fe != null)
            {
                return fe.GetType().Name + (string.IsNullOrEmpty(fe.Name) ? "" : "#" + fe.Name);
            }

            return e.OriginalSource == null ? "null" : e.OriginalSource.GetType().Name;
        }

        #endregion

        #region 下拉框：打开时把当前值滚到中间

        /// <summary>
        /// 下拉框打开时，把当前选中项尽量滚到可视区中间（而非顶/底边缘）。
        /// 默认 ComboBox 只会把选中项带到可视区（通常压在顶部）：
        /// 当年份之类列表很长时，当前值贴在边上，用户要多滚才能看到。
        /// 这里在下拉布局完成后再手动居中一次，覆盖默认行为。
        /// </summary>
        private void CenterSelectedInDropDown(System.Windows.Controls.ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            // 等下拉内容完成布局 / 打开动画后再居中，避免 ViewportHeight 尚未就绪
            combo.Dispatcher.BeginInvoke(
                new Action(() => CenterSelectedInDropDownCore(combo, 0)),
                System.Windows.Threading.DispatcherPriority.Render);
        }

        private void CenterSelectedInDropDownCore(System.Windows.Controls.ComboBox combo, int attempt)
        {
            try
            {
                var popup = combo.Template.FindName("PART_Popup", combo) as System.Windows.Controls.Primitives.Popup;
                if (popup == null || popup.Child == null)
                {
                    return;
                }

                var sv = FindVisualChild<ScrollViewer>(popup.Child);
                if (sv == null)
                {
                    return;
                }

                DependencyObject container = combo.ItemContainerGenerator.ContainerFromItem(combo.SelectedItem);
                if (container == null)
                {
                    container = combo.ItemContainerGenerator.ContainerFromIndex(combo.SelectedIndex);
                }

                var item = container as FrameworkElement;
                if (item == null || sv.ViewportHeight <= 0)
                {
                    // 容器还没生成（虚拟化管理）或视口尚未就绪：下一帧再试，最多 3 次
                    if (attempt < 3)
                    {
                        combo.Dispatcher.BeginInvoke(
                            new Action(() => CenterSelectedInDropDownCore(combo, attempt + 1)),
                            System.Windows.Threading.DispatcherPriority.Render);
                    }
                    return;
                }

                ScrollSelectedToCenter(sv, item);
            }
            catch
            {
                // 居中失败不致命，下拉仍可正常使用
            }
        }

        /// <summary>
        /// 把 item 滚到 ScrollViewer 可视区垂直居中位置（超出范围时夹紧到首/尾）。
        /// </summary>
        private static void ScrollSelectedToCenter(ScrollViewer sv, FrameworkElement item)
        {
            try
            {
                var topLeft = item.TransformToVisual(sv).Transform(new Point(0, 0));
                double offset = topLeft.Y + (item.ActualHeight - sv.ViewportHeight) / 2.0;
                if (offset < 0) { offset = 0; }
                if (offset > sv.ScrollableHeight) { offset = sv.ScrollableHeight; }
                sv.ScrollToVerticalOffset(offset);
            }
            catch
            {
                // ignore
            }
        }

        /// <summary>
        /// 窗口隧道阶段的滚轮接管：任一下拉开着时优先滚动弹层内容；
        /// 若列表一屏放得下（滑动窗口仅 11 项，ScrollableHeight=0、无可滚内容，
        /// 这正是“滚轮滚不动”的真正原因），则把滚轮转成「切换选中值」——
        /// 向上=上一个、向下=下一个，选中变化经 Cmb_SelectionChanged 以新值为锚
        /// 重建窗口（选中恒居中），于是滚轮可连续转到任意年/月/日。
        /// </summary>
        private void Win_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            try
            {
                var combo = CmbYear.IsDropDownOpen ? CmbYear
                          : CmbMonth.IsDropDownOpen ? CmbMonth
                          : CmbDay.IsDropDownOpen ? CmbDay
                          : null;
                if (combo == null)
                {
                    return; // 没有下拉开着：不干预默认行为（日历等区域照常）
                }

                // 列表超出弹层高度（小屏 MaxDropDownHeight 受限）时：滚内容
                var sv = FindDropDownScrollViewer(combo);
                if (sv != null && sv.ScrollableHeight > 0)
                {
                    // 滚轮向上（Delta>0）→ 滚动偏移减小
                    sv.ScrollToVerticalOffset(sv.VerticalOffset - e.Delta);
                    e.Handled = true;
                    Trace("wheel scroll " + combo.Name + " offset=" + sv.VerticalOffset.ToString("F0") + " delta=" + e.Delta);
                    return;
                }

                // 一屏放得下、无可滚内容：滚轮切换选中值
                int idx = combo.SelectedIndex;
                if (idx < 0)
                {
                    return;
                }

                int next = idx + (e.Delta > 0 ? -1 : 1);
                if (next < 0 || next >= combo.Items.Count)
                {
                    return;
                }

                combo.SelectedIndex = next;
                e.Handled = true;
                Trace("wheel change " + combo.Name + " -> " + (combo.SelectedItem ?? "null") + " delta=" + e.Delta);
            }
            catch
            {
                // 滚轮接管失败不致命：退回默认行为
            }
        }

        /// <summary>
        /// 取下拉弹层内部的 ScrollViewer（HandyControl 模板：PART_Popup → Border → ScrollViewer）。
        /// </summary>
        private ScrollViewer FindDropDownScrollViewer(System.Windows.Controls.ComboBox combo)
        {
            try
            {
                var popup = combo.Template?.FindName("PART_Popup", combo)
                            as System.Windows.Controls.Primitives.Popup;
                if (popup == null || popup.Child == null)
                {
                    return null;
                }

                return FindVisualChild<ScrollViewer>(popup.Child);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 下拉弹层提速：HandyControl 的 ComboBox 模板给 Popup 配了
        /// PopupAnimation=Fade（系统级渐显，约 200ms）和 DropShadowEffect 投影
        /// （分层窗口上逐帧模糊渲染），两者叠加即“下拉反应慢”。
        /// 在模板加载后（Loaded，弹层尚未打开）直接移除，弹层即时出现。
        /// </summary>
        private static void SpeedUpDropDown(System.Windows.Controls.ComboBox combo)
        {
            try
            {
                var popup = combo.Template?.FindName("PART_Popup", combo)
                            as System.Windows.Controls.Primitives.Popup;
                if (popup == null)
                {
                    return; // 模板非 HandyControl（回退系统样式）：无需处理
                }

                popup.PopupAnimation = System.Windows.Controls.Primitives.PopupAnimation.None;

                if (popup.Child is Border border)
                {
                    border.Effect = null;              // 去投影（原 Margin 是给投影留的余量）
                    border.Margin = new Thickness(0);  // 弹层与下拉框边缘对齐，不留投影空隙
                }
            }
            catch
            {
                // 提速失败不影响下拉功能
            }
        }

        private static T FindVisualChild<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null)
            {
                return null;
            }

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T t)
                {
                    return t;
                }

                var found = FindVisualChild<T>(child);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        #endregion

        /// <summary>
        /// 加载标题栏图标。
        /// 注意：HandControl 的 NonClientAreaContent 会替换系统标题栏内容，
        /// 这里只给自绘标题栏的大图标赋值；系统标题栏的小图标不设置（保持干净）。
        /// </summary>
        private void LoadTitleIcon()
        {
            try
            {
                var asm = typeof(DatePickerWindow).Assembly;
                using (var s = asm.GetManifestResourceStream("OutlookDateJump.CalendarIcon.png"))
                {
                    if (s == null)
                    {
                        return;
                    }

                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = s;
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    ImgTitleIcon.Source = bmp;
                }
            }
            catch
            {
                // title icon is non-critical
            }
        }

        /// <summary>
        /// 拖动自绘标题栏移动窗口。
        /// </summary>
        private void TitleBar_Drag(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch
            {
                // ignore drag failures
            }
        }

        private void InitCombos(DateTime initial)
        {
            // 滑动窗口：三个下拉都只显示「锚点值 ±WindowSpan」，选中项恒在中间。
            // 先重建窗口再选中——否则目标值不在列表里，赋值会变成 null。
            _busy = true;
            try
            {
                FillYears(initial.Year);
                FillMonths(initial.Month);
                FillDayWindow(initial.Year, initial.Month, initial.Day);

                CmbYear.SelectedItem = initial.Year;
                CmbMonth.SelectedItem = initial.Month;
                CmbDay.SelectedItem = initial.Day;
            }
            finally
            {
                _busy = false;
            }

            Cal.DisplayDate = initial;
            Cal.SelectedDate = initial;
        }

        private void SetDate(DateTime dt)
        {
            // 与 InitCombos 同理：先按目标日期重建三个窗口，再选中。
            // 「今天」按钮 / 日历点选都走这里，目标日期可能远在当前窗口之外。
            _busy = true;
            try
            {
                FillYears(dt.Year);
                FillMonths(dt.Month);
                FillDayWindow(dt.Year, dt.Month, dt.Day);

                CmbYear.SelectedItem = dt.Year;
                CmbMonth.SelectedItem = dt.Month;
                CmbDay.SelectedItem = dt.Day;
            }
            finally
            {
                _busy = false;
            }

            Cal.DisplayDate = dt;
            Cal.SelectedDate = dt;
        }

        /// <summary>
        /// 滑动窗口半宽：年/月/日下拉都显示「锚点值前后各 5 个」。
        /// 点到边界项后列表以新选中值为锚点重建，于是一步步能滑到任意年月日。
        /// </summary>
        private const int WindowSpan = 5;

        /// <summary>以 anchor 为中心重建「年」下拉窗口（11 项，选中 anchor）。</summary>
        private void FillYears(int anchor)
        {
            CmbYear.Items.Clear();
            for (int i = anchor - WindowSpan; i <= anchor + WindowSpan; i++)
            {
                CmbYear.Items.Add(i);
            }
            CmbYear.SelectedItem = anchor;
        }

        /// <summary>
        /// 以 anchor 为中心重建「月」下拉窗口（11 项，跨年环绕）。
        /// 例：anchor=10 → 5,6,7,8,9,10,11,12,1,2,3（10 在中间）。
        /// </summary>
        private void FillMonths(int anchor)
        {
            CmbMonth.Items.Clear();
            for (int i = -WindowSpan; i <= WindowSpan; i++)
            {
                CmbMonth.Items.Add(WrapMonth(anchor + i));
            }
            CmbMonth.SelectedItem = anchor;
        }

        /// <summary>把任意月份值绕回 1–12。</summary>
        private static int WrapMonth(int m)
        {
            return ((m - 1) % 12 + 12) % 12 + 1;
        }

        /// <summary>
        /// 以 anchorDay 为中心重建「日」下拉窗口：取当月 1–月末 与
        /// [anchorDay-5, anchorDay+5] 的交集（靠近月初/月末时不足 11 项），
        /// anchorDay 超出当月天数时收拢到月末。选中 anchor。
        /// </summary>
        private void FillDayWindow(int year, int month, int anchorDay)
        {
            int days = DateTime.DaysInMonth(year, month);
            int anchor = Math.Min(Math.Max(anchorDay, 1), days);

            CmbDay.Items.Clear();
            int lo = Math.Max(1, anchor - WindowSpan);
            int hi = Math.Min(days, anchor + WindowSpan);
            for (int d = lo; d <= hi; d++)
            {
                CmbDay.Items.Add(d);
            }
            CmbDay.SelectedItem = anchor;
        }

        private void Cmb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_busy)
            {
                return;
            }

            if (CmbYear.SelectedItem == null || CmbMonth.SelectedItem == null)
            {
                return;
            }

            _busy = true;
            try
            {
                int y = (int)CmbYear.SelectedItem;
                int m = (int)CmbMonth.SelectedItem;

                // 哪个下拉变了，就以其新选中值为锚点重建哪个窗口（选中项恒居中）；
                // 点到边界项即重新锚定，于是一步步能滑到任意年月日。
                if (sender == CmbYear)
                {
                    FillYears(y);
                }
                else if (sender == CmbMonth)
                {
                    FillMonths(m);
                }

                // 关键：年/月变化时先重算「日」窗口（如 31 日切到 2 月收拢为月末），
                // 否则 new DateTime 会抛异常被吞掉，Cal.SelectedDate 保持旧值，
                // 点「跳转」就会跳回原日期——表现为“第一次没反应”。
                int curDay = (CmbDay.SelectedItem is int d0) ? d0 : DateTime.DaysInMonth(y, m);
                FillDayWindow(y, m, curDay);

                if (CmbDay.SelectedItem == null)
                {
                    return;
                }

                int d = (int)CmbDay.SelectedItem;
                var dt = new DateTime(y, m, d);
                Cal.DisplayDate = dt;
                Cal.SelectedDate = dt;
            }
            catch
            {
                // ignore invalid combinations
            }
            finally
            {
                _busy = false;
            }
        }

        private void Cal_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Cal.SelectedDate == null)
            {
                return;
            }

            // 由 Cmb_SelectionChanged 设置 Cal.SelectedDate 引起的嵌套调用直接忽略：
            // 否则会把外层尚在使用的 _busy 提前释放，且重建/日志全是重复劳动。
            if (_busy)
            {
                return;
            }

            var dt = Cal.SelectedDate.Value;
            Trace("Cal_SelectedDatesChanged dt=" + dt.ToString("yyyy-MM-dd") +
                  " CmbDay.Item=" + (CmbDay.SelectedItem ?? "null"));

            _busy = true;
            try
            {
                // 先按日历选中的日期重建三个窗口，否则目标值不在列表里，赋值会变 null
                FillYears(dt.Year);
                FillMonths(dt.Month);
                FillDayWindow(dt.Year, dt.Month, dt.Day);

                CmbYear.SelectedItem = dt.Year;
                CmbMonth.SelectedItem = dt.Month;
                CmbDay.SelectedItem = dt.Day;
            }
            finally
            {
                _busy = false;
            }

            // 关键：点完日历后清理“上一次交互”的残留状态，
            // 否则用户接下来对「跳转」按钮的第一次点击会被这些残留吃掉
            // （表现为要点两次才跳转）：
            //   1) 日历的日期按钮会捕获鼠标，捕获不释放则下一次点击被路由给日历；
            //   2) 下拉框的弹层若还认为自己开着，第一次点击只会用来关弹层。
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    Trace("cleanup: capture=" + CaptureName() +
                          " dropDown=" + (CmbYear.IsDropDownOpen ? 1 : 0) + (CmbMonth.IsDropDownOpen ? 1 : 0) + (CmbDay.IsDropDownOpen ? 1 : 0));

                    if (CmbYear.IsDropDownOpen) { CmbYear.IsDropDownOpen = false; }
                    if (CmbMonth.IsDropDownOpen) { CmbMonth.IsDropDownOpen = false; }
                    if (CmbDay.IsDropDownOpen) { CmbDay.IsDropDownOpen = false; }

                    if (System.Windows.Input.Mouse.Captured != null)
                    {
                        System.Windows.Input.Mouse.Capture(null);
                    }
                }
                catch { }
            }), DispatcherPriority.Input);

            Trace("  after sync: CmbDay.Item=" + (CmbDay.SelectedItem ?? "null") +
                  " Cal.SelectedDate=" + (Cal.SelectedDate.HasValue ? Cal.SelectedDate.Value.ToString("yyyy-MM-dd") : "null"));
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

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            ConfirmJump();
        }

        /// <summary>
        /// 确认跳转。独立方法，供「跳转」按钮的 MouseDown / MouseUp / Click 共用。
        /// 只负责记录选中日期并关闭窗口——不做任何 Outlook COM 调用，
        /// 保证点击后窗口立即关闭（跳转由宿主在窗口关闭后调度）。
        /// </summary>
        private void ConfirmJump()
        {
            if (_confirmed)
            {
                return;
            }
            _confirmed = true;

            Trace("BtnOk Cal.SelectedDate=" + (Cal.SelectedDate.HasValue ? Cal.SelectedDate.Value.ToString("yyyy-MM-dd") : "NULL") +
                  " CmbDay=" + (CmbDay.SelectedItem ?? "null"));
            SelectedDate = Cal.SelectedDate ?? DateTime.Today;

            try
            {
                DialogResult = true;
            }
            catch
            {
                // 若窗口已关闭则忽略
            }
        }

        /// <summary>
        /// 双击日历某一天 = 直接确认跳转，无需再点「跳转」按钮。
        /// 注意：只认落在「日期格子」(CalendarDayButton) 上的双击；
        /// 落在表头、月份左右调节箭头、空白区域的不算，否则在箭头上快速连点
        /// （被 Windows 识别成一次双击）会误关窗。
        /// </summary>
        private void Cal_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (Cal.SelectedDate == null || _confirmed)
            {
                return;
            }

            try
            {
                var hit = Cal.InputHitTest(e.GetPosition(Cal)) as DependencyObject;
                if (!IsSelfOrDescendantOfType(hit, typeof(System.Windows.Controls.Primitives.CalendarDayButton)))
                {
                    Trace("DblClick ignored (not on a day cell)");
                    return;
                }
            }
            catch
            {
                // 命中测试异常时退化为"不确认"，避免误关窗
                return;
            }

            _confirmed = true;
            Trace("DblClick " + Cal.SelectedDate.Value.ToString("yyyy-MM-dd"));
            SelectedDate = Cal.SelectedDate;
            DialogResult = true;
        }

        private void BtnToday_Click(object sender, RoutedEventArgs e)
        {
            SetDate(DateTime.Today);
        }

        /// <summary>
        /// 「今天」按钮：鼠标按下即响应（与「跳转」一致，规避首次 Click 丢失）。
        /// </summary>
        private void BtnToday_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            SetDate(DateTime.Today);
        }
    }
}
