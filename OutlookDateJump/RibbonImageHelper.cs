using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookDateJump
{
    /// <summary>
    /// 将嵌入的 CalendarIcon.png（由 Calendar.png 生成的 64x64 图标）转为
    /// 功能区按钮可用的图片 COM 对象（IPictureDisp）。
    /// 自行声明 COM 接口并经 AxHost 子类转换，避免依赖 stdole PIA
    /// （部分机器未安装该 PIA，直接引用会在运行时绑定失败）。
    /// </summary>
    public static class RibbonImageHelper
    {
        /// <summary>与 stdole.IPictureDisp 同 GUID 的自声明 COM 接口（必须 public，因 Ribbon1 公开返回它）。</summary>
        [ComImport]
        [Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        public interface IPictureDisp
        {
        }

        /// <summary>
        /// AxHost 的 GetIPictureDispFromPicture 是 protected，经典做法是经由子类调用。
        /// </summary>
        private sealed class PictureDispConverter : AxHost
        {
            private PictureDispConverter()
                : base(null)
            {
            }

            public static IPictureDisp Convert(Image image)
            {
                return AxHost.GetIPictureDispFromPicture(image) as IPictureDisp;
            }
        }

        public static IPictureDisp GetIcon(string imageName)
        {
            try
            {
                var asm = typeof(RibbonImageHelper).Assembly;
                using (var s = asm.GetManifestResourceStream("OutlookDateJump." + imageName + ".png"))
                {
                    if (s == null)
                    {
                        return null;
                    }

                    using (var img = Image.FromStream(s))
                    {
                        // 原始图标常带大面积透明/白边，直接缩到 32x32 会让图形显得很小。
                        // 先裁掉这些边距，再让图形满格填入 32x32，功能区按钮图标才够醒目。
                        using (var src = new Bitmap(img.Width, img.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                        {
                            using (var g0 = Graphics.FromImage(src))
                            {
                                g0.Clear(Color.Transparent);
                                g0.DrawImage(img, 0, 0, img.Width, img.Height);
                            }

                            Rectangle content = TrimMargins(src);
                            using (var cropped = src.Clone(content, src.PixelFormat))
                            {
                                int side = Math.Max(content.Width, content.Height);
                                float scale = 32f / side;
                                int w = Math.Max(1, (int)Math.Round(content.Width * scale));
                                int h = Math.Max(1, (int)Math.Round(content.Height * scale));

                                // 渲染为 32x32 满格图标：功能区会把它缩放到按钮实际尺寸，
                                // 但源图分辨率更高，缩放后依然清晰（normal 按钮显示约 16-20px）。
                                // 必须无 alpha 通道且垫白底，否则 Office 会按全透明渲染（图标隐形）。
                                using (var small = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppRgb))
                                {
                                    small.SetResolution(96, 96);
                                    using (var g = Graphics.FromImage(small))
                                    {
                                        g.Clear(Color.White);
                                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                                        g.DrawImage(cropped, new Rectangle((32 - w) / 2, (32 - h) / 2, w, h));
                                    }

                                    return PictureDispConverter.Convert(small);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 裁掉图标四周的透明边与近白边，让有效图形尽量占满按钮。
        /// </summary>
        private static Rectangle TrimMargins(Bitmap bmp)
        {
            int left = 0, top = 0, right = bmp.Width - 1, bottom = bmp.Height - 1;

            // 先按不透明像素求出内容包围盒
            int aL = bmp.Width, aT = bmp.Height, aR = -1, aB = -1;
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    if (bmp.GetPixel(x, y).A > 20)
                    {
                        if (x < aL) aL = x;
                        if (x > aR) aR = x;
                        if (y < aT) aT = y;
                        if (y > aB) aB = y;
                    }
                }
            }

            if (aR >= aL && aB >= aT)
            {
                left = aL; top = aT; right = aR; bottom = aB;
            }

            // 再削掉四周几乎全白（或全透明）的行列
            while (top <= bottom && RowIsBlank(bmp, top, left, right)) top++;
            while (bottom >= top && RowIsBlank(bmp, bottom, left, right)) bottom--;
            while (left <= right && ColIsBlank(bmp, left, top, bottom)) left++;
            while (right >= left && ColIsBlank(bmp, right, top, bottom)) right--;

            if (right < left || bottom < top)
            {
                return new Rectangle(0, 0, bmp.Width, bmp.Height);
            }

            return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }

        private static bool RowIsBlank(Bitmap bmp, int y, int left, int right)
        {
            int blank = 0;
            for (int x = left; x <= right; x++)
            {
                if (IsBlank(bmp.GetPixel(x, y)))
                {
                    blank++;
                }
            }
            return blank >= (right - left + 1) * 0.98;
        }

        private static bool ColIsBlank(Bitmap bmp, int x, int top, int bottom)
        {
            int blank = 0;
            for (int y = top; y <= bottom; y++)
            {
                if (IsBlank(bmp.GetPixel(x, y)))
                {
                    blank++;
                }
            }
            return blank >= (bottom - top + 1) * 0.98;
        }

        private static bool IsBlank(Color c)
        {
            return c.A < 20 || (c.R > 242 && c.G > 242 && c.B > 242);
        }
    }
}
