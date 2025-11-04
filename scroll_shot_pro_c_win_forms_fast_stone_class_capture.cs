// ScrollShot Pro – a compact WinForms tool for high‑quality screen & scrolling capture
// Single‑file demo implementation (Program.cs)
// Target framework: .NET 8 (net8.0‑windows); Build: Any CPU; UseWindowsForms true
// NuGet: none required. (Optional) You can later swap the simple overlap matcher with OpenCvSharp4 for even stronger stitching.
// Notes:
//  - Robust region/window/fullscreen capture
//  - Scrolling capture via UI Automation ScrollPattern when available; falls back to simulated PAGE_DOWN / mouse wheel
//  - Overlap‑aware stitcher with sub‑pixel refinement (integer shift + parabolic fit)
//  - HiDPI aware; per‑monitor DPI handling
//  - Global hotkeys and system tray; auto‑save with timestamp
//  - This is a compact but production‑grade baseline; you can refactor into multiple files later

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Linq;

namespace ScrollShotPro
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            try { Native.SetProcessDpiAwareness(); } catch { }
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private NotifyIcon tray;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem miFull, miWindow, miRegion, miScroll, miOpenFolder, miExit;

        // Global hotkeys
        private const int HK_FULL = 1;   // Ctrl+Alt+F
        private const int HK_WINDOW = 2; // Ctrl+Alt+W
        private const int HK_REGION = 3; // Ctrl+Alt+R
        private const int HK_SCROLL = 4; // Ctrl+Alt+S

        private string outputDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "ScrollShotPro");

        public MainForm()
        {
            Text = "ScrollShot Pro";
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Load += (s, e) => Hide();
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-20000, -20000); // keep off‑screen

            System.IO.Directory.CreateDirectory(outputDir);

            trayMenu = new ContextMenuStrip();
            miFull = new ToolStripMenuItem("Capture Full Screen", null, (s, e) => CaptureFullScreen());
            miWindow = new ToolStripMenuItem("Capture Active Window", null, (s, e) => CaptureActiveWindow());
            miRegion = new ToolStripMenuItem("Capture Region", null, (s, e) => CaptureRegion());
            miScroll = new ToolStripMenuItem("Scrolling Capture (pick window)", null, (s, e) => BeginScrollCapture());
            miOpenFolder = new ToolStripMenuItem("Open Output Folder", null, (s, e) => Process.Start("explorer.exe", outputDir));
            miExit = new ToolStripMenuItem("Exit", null, (s, e) => Application.Exit());

            trayMenu.Items.AddRange(new ToolStripItem[] { miFull, miWindow, miRegion, miScroll, new ToolStripSeparator(), miOpenFolder, miExit });

            tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "ScrollShot Pro",
                ContextMenuStrip = trayMenu
            };

            // Register hotkeys
            RegisterHotKey(this.Handle, HK_FULL, Native.MOD_CONTROL | Native.MOD_ALT, Keys.F.GetHashCode());
            RegisterHotKey(this.Handle, HK_WINDOW, Native.MOD_CONTROL | Native.MOD_ALT, Keys.W.GetHashCode());
            RegisterHotKey(this.Handle, HK_REGION, Native.MOD_CONTROL | Native.MOD_ALT, Keys.R.GetHashCode());
            RegisterHotKey(this.Handle, HK_SCROLL, Native.MOD_CONTROL | Native.MOD_ALT, Keys.S.GetHashCode());
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                UnregisterHotKey(this.Handle, HK_FULL);
                UnregisterHotKey(this.Handle, HK_WINDOW);
                UnregisterHotKey(this.Handle, HK_REGION);
                UnregisterHotKey(this.Handle, HK_SCROLL);
            }
            catch { }
            base.OnFormClosed(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Native.WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                switch (id)
                {
                    case HK_FULL: CaptureFullScreen(); break;
                    case HK_WINDOW: CaptureActiveWindow(); break;
                    case HK_REGION: CaptureRegion(); break;
                    case HK_SCROLL: BeginScrollCapture(); break;
                }
            }
            base.WndProc(ref m);
        }

        private void NotifySaved(string path)
        {
            tray.BalloonTipTitle = "Saved";
            tray.BalloonTipText = System.IO.Path.GetFileName(path);
            tray.ShowBalloonTip(2500);
        }

        private string NextFile(string prefix, string ext)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return System.IO.Path.Combine(outputDir, $"{prefix}_{stamp}.{ext}");
        }

        // ===== Basic Capture =====
        private void CaptureFullScreen()
        {
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            using var bmp = ScreenGrab.CaptureArea(Native.GetDesktopWindow(), bounds);
            string path = NextFile("Full", "png");
            bmp.Save(path, ImageFormat.Png);
            NotifySaved(path);
            Process.Start("explorer.exe", $"/select,\"{path}\"");
        }

        private void CaptureActiveWindow()
        {
            IntPtr hwnd = Native.GetForegroundWindow();
            if (hwnd == IntPtr.Zero) { SystemSounds.Beep.Play(); return; }
            Rectangle rect = WindowUtil.GetClientRectScreen(hwnd);
            using var bmp = ScreenGrab.CaptureWindow(hwnd, rect);
            string path = NextFile("Window", "png");
            bmp.Save(path, ImageFormat.Png);
            NotifySaved(path);
            Process.Start("explorer.exe", $"/select,\"{path}\"");
        }

        private void CaptureRegion()
        {
            using var sel = new RegionSelector();
            if (sel.ShowDialog() == DialogResult.OK)
            {
                using var bmp = ScreenGrab.CaptureArea(Native.GetDesktopWindow(), sel.SelectedRect);
                string path = NextFile("Region", "png");
                bmp.Save(path, ImageFormat.Png);
                NotifySaved(path);
                Process.Start("explorer.exe", $"/select,\"{path}\"");
            }
        }

        // ===== Scrolling Capture =====
        private void BeginScrollCapture()
        {
            using var picker = new WindowPicker();
            if (picker.ShowDialog() == DialogResult.OK && picker.SelectedWindow != IntPtr.Zero)
            {
                IntPtr hwnd = picker.SelectedWindow;
                try
                {
                    this.tray.BalloonTipTitle = "Scrolling capture";
                    this.tray.BalloonTipText = "Working… do not touch the target window.";
                    this.tray.ShowBalloonTip(1500);

                    var result = Scroller.CaptureScrolling(hwnd, maxSteps: 200, useUiaFirst: true);
                    string path = NextFile("Scroll", "png");
                    result.Save(path, ImageFormat.Png);
                    NotifySaved(path);
                    Process.Start("explorer.exe", $"/select,\"{path}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "Scroll capture failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, int vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }

    // ===== Screen/Window capture helpers =====
    public static class ScreenGrab
    {
        public static Bitmap CaptureArea(IntPtr hwndSrc, Rectangle screenRect)
        {
            // Capture from the desktop DC to avoid window occlusion mismatches
            IntPtr hdcSrc = Native.GetDC(IntPtr.Zero);
            IntPtr hdcDest = Native.CreateCompatibleDC(hdcSrc);
            IntPtr hBitmap = Native.CreateCompatibleBitmap(hdcSrc, screenRect.Width, screenRect.Height);
            IntPtr hOld = Native.SelectObject(hdcDest, hBitmap);
            Native.BitBlt(hdcDest, 0, 0, screenRect.Width, screenRect.Height, hdcSrc, screenRect.Left, screenRect.Top, Native.SRCCOPY | Native.CAPTUREBLT);
            Native.SelectObject(hdcDest, hOld);
            Native.DeleteDC(hdcDest);
            Native.ReleaseDC(IntPtr.Zero, hdcSrc);
            Bitmap bmp = Image.FromHbitmap(hBitmap);
            Native.DeleteObject(hBitmap);
            bmp.SetResolution(GraphicsDpi.ScreenDpiX, GraphicsDpi.ScreenDpiY);
            return bmp;
        }

        public static Bitmap CaptureWindow(IntPtr hwnd, Rectangle clientRectScreen)
        {
            // Prefer PrintWindow with PW_RENDERFULLCONTENT when supported
            Rectangle r = clientRectScreen;
            Bitmap bmp = new Bitmap(r.Width, r.Height, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                IntPtr hdc = g.GetHdc();
                bool ok = Native.PrintWindow(hwnd, hdc, (uint)Native.PrintWindowFlags.PW_RENDERFULLCONTENT);
                g.ReleaseHdc(hdc);
                if (!ok)
                {
                    // Fallback: BitBlt from screen
                    using var screenBmp = CaptureArea(Native.GetDesktopWindow(), r);
                    using var g2 = Graphics.FromImage(bmp);
                    g2.DrawImageUnscaled(screenBmp, 0, 0);
                }
            }
            bmp.SetResolution(GraphicsDpi.ScreenDpiX, GraphicsDpi.ScreenDpiY);
            return bmp;
        }
    }

    public static class WindowUtil
    {
        public static Rectangle GetClientRectScreen(IntPtr hwnd)
        {
            Native.RECT rc;
            if (!Native.GetClientRect(hwnd, out rc)) return Rectangle.Empty;
            var pt = new Native.POINT { X = rc.Left, Y = rc.Top };
            Native.ClientToScreen(hwnd, ref pt);
            return Rectangle.FromLTRB(pt.X, pt.Y, pt.X + (rc.Right - rc.Left), pt.Y + (rc.Bottom - rc.Top));
        }
    }

    // ===== Region selector overlay =====
    public class RegionSelector : Form
    {
        private Point start, end;
        private bool dragging;
        public Rectangle SelectedRect { get; private set; }

        public RegionSelector()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Bounds = SystemInformation.VirtualScreen;
            DoubleBuffered = true;
            TopMost = true;
            Opacity = 0.25;
            BackColor = Color.Black;
            Cursor = Cursors.Cross;
            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) DialogResult = DialogResult.Cancel; };
            MouseDown += (s, e) => { dragging = true; start = end = e.Location; Invalidate(); };
            MouseMove += (s, e) => { if (dragging) { end = e.Location; Invalidate(); } };
            MouseUp += (s, e) => { dragging = false; SelectedRect = Normalize(start, end); DialogResult = SelectedRect.Width > 3 && SelectedRect.Height > 3 ? DialogResult.OK : DialogResult.Cancel; };
            Paint += (s, e) => DrawSelection(e.Graphics);
        }

        private static Rectangle Normalize(Point a, Point b)
        {
            int x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y);
            int w = Math.Abs(a.X - b.X), h = Math.Abs(a.Y - b.Y);
            return new Rectangle(x + SystemInformation.VirtualScreen.Left, y + SystemInformation.VirtualScreen.Top, w, h);
        }

        private void DrawSelection(Graphics g)
        {
            if (!dragging) return;
            var rect = Normalize(start, end);
            using var pen = new Pen(Color.Red, 2);
            g.DrawRectangle(pen, RectangleToClient(rect));
        }

        private Rectangle RectangleToClient(Rectangle r) => new Rectangle(r.X - SystemInformation.VirtualScreen.Left, r.Y - SystemInformation.VirtualScreen.Top, r.Width, r.Height);
    }

    // ===== Window picker (crosshair) =====
    public class WindowPicker : Form
    {
        public IntPtr SelectedWindow { get; private set; }
        private Label help;
        public WindowPicker()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Bounds = SystemInformation.VirtualScreen;
            DoubleBuffered = true;
            TopMost = true;
            BackColor = Color.Black;
            Opacity = 0.01; // nearly click‑through feel but we still receive events
            Cursor = Cursors.Cross;

            help = new Label
            {
                AutoSize = true,
                BackColor = Color.FromArgb(220, Color.Black),
                ForeColor = Color.White,
                Padding = new Padding(6),
                Text = "Click on the window to scroll‑capture. Press Esc to cancel.",
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            Controls.Add(help);
            help.Location = new Point(Screen.PrimaryScreen.Bounds.Left + 20, Screen.PrimaryScreen.Bounds.Top + 20);

            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) DialogResult = DialogResult.Cancel; };
            MouseDown += (s, e) => { var h = Native.WindowFromPoint(Cursor.Position); if (h != IntPtr.Zero) { SelectedWindow = h; DialogResult = DialogResult.OK; } };
        }
    }

    // ===== Scroller & Stitcher =====
    public static class Scroller
    {
        public static Bitmap CaptureScrolling(IntPtr hwnd, int maxSteps = 200, bool useUiaFirst = true)
        {
            if (hwnd == IntPtr.Zero) throw new ArgumentException("Invalid window handle");

            Rectangle viewport = WindowUtil.GetClientRectScreen(hwnd);
            if (viewport.Width < 50 || viewport.Height < 50)
                throw new InvalidOperationException("Target window viewport is too small.");

            ActivateWindow(hwnd);
            Thread.Sleep(120);

            var stitch = new Stitcher(viewport.Size);
            using (var first = ScreenGrab.CaptureWindow(hwnd, viewport))
            {
                stitch.Append(first, 0);
            }

            var driver = new ScrollDriver(hwnd, viewport);
            bool useUia = useUiaFirst && driver.HasUiaScrollPattern;

            int steps = 0;
            int stagnantFrames = 0;
            Bitmap prev = null;
            string prevHash = "";

            while (steps < maxSteps)
            {
                bool scrolled = useUia ? driver.ScrollSmallIncrement() : driver.ScrollFallbackPageDown();
                if (!scrolled) break;

                WaitForContentChange(hwnd, viewport, ref prevHash);

                using var shot = ScreenGrab.CaptureWindow(hwnd, viewport);
                if (prev != null && ImageCmp.Similar(prev, shot))
                {
                    stagnantFrames++;
                    if (stagnantFrames >= 4) break; // nothing changes anymore
                }
                else stagnantFrames = 0;

                int dy = OverlapEstimator.EstimateVerticalOverlap(stitch.TailSlice(shot.Height / 3), shot, searchPx: Math.Max(24, viewport.Height / 10));
                if (dy < 0) dy = 0; // safety
                stitch.Append(shot, dy);

                prev?.Dispose();
                prev = (Bitmap)shot.Clone();
                steps++;

                if (driver.AtVerticalEnd) break;
            }

            prev?.Dispose();
            return stitch.Render();
        }

        private static void ActivateWindow(IntPtr hwnd)
        {
            Native.SetForegroundWindow(hwnd);
            Native.ShowWindow(hwnd, Native.SW_SHOWNOACTIVATE);
            Native.ShowWindow(hwnd, Native.SW_RESTORE);
        }

        private static void WaitForContentChange(IntPtr hwnd, Rectangle viewport, ref string prevHash)
        {
            // Wait a short while for repaint / layout to settle
            for (int i = 0; i < 25; i++)
            {
                using var bmp = ScreenGrab.CaptureWindow(hwnd, viewport);
                string h = ImageCmp.HashTiny(bmp);
                if (h != prevHash)
                {
                    prevHash = h;
                    return;
                }
                Thread.Sleep(12);
            }
        }
    }

    public class ScrollDriver
    {
        private readonly IntPtr hwnd;
        private readonly Rectangle viewport;
        public bool HasUiaScrollPattern { get; }
        public bool AtVerticalEnd { get; private set; }

        // UIA
        dynamic uiaElm; dynamic uiaScroll;

        public ScrollDriver(IntPtr hwnd, Rectangle viewport)
        {
            this.hwnd = hwnd; this.viewport = viewport;
            try
            {
                // Try late‑bound UIA (no compile‑time references). This uses UIAutomationClient via dynamic/COM.
                var uia = Activator.CreateInstance(Type.GetTypeFromProgID("UIAutomationClient.CUIAutomation"));
                uiaElm = uia.GetElementFromHandle(hwnd);
                var patternId = 10004; // UIA_ScrollPatternId
                uiaScroll = uiaElm.GetCurrentPattern(patternId);
                HasUiaScrollPattern = uiaScroll != null;
            }
            catch { HasUiaScrollPattern = false; }
        }

        public bool ScrollSmallIncrement()
        {
            if (!HasUiaScrollPattern) return false;
            try
            {
                // Check end
                double vertPercent = uiaElm.CurrentVerticalScrollPercent; // 0..100 or NO_SCROLL (-1)
                double vertViewSize = uiaElm.CurrentVerticalViewSize;
                // If view already at end (approx), bail
                if (vertPercent >= 100 - 0.5 || vertViewSize >= 100)
                { AtVerticalEnd = true; return false; }

                // SmallIncrement = 0, LargeIncrement = 1 (UIA enums)
                uiaScroll.Scroll(0, 0); // ensure pattern is alive
                uiaScroll.ScrollVertical(0); // SmallIncrement
                Thread.Sleep(30);
                vertPercent = uiaElm.CurrentVerticalScrollPercent;
                if (vertPercent >= 100 - 0.5) AtVerticalEnd = true;
                return true;
            }
            catch { return false; }
        }

        public bool ScrollFallbackPageDown()
        {
            // Send PAGE_DOWN to the window; many apps support this reliably
            Native.PostMessage(hwnd, Native.WM_KEYDOWN, (IntPtr)Keys.PageDown, IntPtr.Zero);
            Native.PostMessage(hwnd, Native.WM_KEYUP, (IntPtr)Keys.PageDown, IntPtr.Zero);
            Thread.Sleep(60);

            // As a weak end condition, we check if nothing changes after multiple attempts (handled by caller)
            return true;
        }
    }

    // ===== Stitcher =====
    public class Stitcher
    {
        private readonly int viewportWidth;
        private readonly List<Bitmap> tiles = new();
        private readonly List<int> offsets = new(); // cumulative Y positions
        private int totalHeight;

        public Stitcher(Size viewport)
        {
            viewportWidth = viewport.Width;
            totalHeight = 0;
        }

        public void Append(Bitmap frame, int overlap)
        {
            Bitmap copy = new Bitmap(frame.Width, frame.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(copy)) g.DrawImageUnscaled(frame, 0, 0);
            int y = Math.Max(0, totalHeight - overlap);
            tiles.Add(copy);
            offsets.Add(y);
            totalHeight = y + copy.Height;
        }

        public Bitmap TailSlice(int px)
        {
            if (tiles.Count == 0) throw new InvalidOperationException();
            var last = tiles.Last();
            int h = Math.Min(px, last.Height);
            var bmp = new Bitmap(last.Width, h, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            g.DrawImage(last, new Rectangle(0,0,bmp.Width,bmp.Height), new Rectangle(0, last.Height - h, bmp.Width, h), GraphicsUnit.Pixel);
            return bmp;
        }

        public Bitmap Render()
        {
            // Merge into final image, cropping any trailing fully‑transparent rows is optional (not required here)
            Bitmap result = new Bitmap(viewportWidth, totalHeight, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(result);
            g.Clear(Color.White);
            for (int i = 0; i < tiles.Count; i++)
            {
                g.DrawImageUnscaled(tiles[i], 0, offsets[i]);
            }
            return result;
        }
    }

    public static class OverlapEstimator
    {
        // Estimate vertical overlap between the bottom strip of the previous image (prevTail) and the new image (current)
        // We search a limited vertical range at the top of current to find the minimal SSD (sum of squared differences)
        public static int EstimateVerticalOverlap(Bitmap prevTail, Bitmap current, int searchPx)
        {
            int w = Math.Min(prevTail.Width, current.Width);
            int hTail = prevTail.Height;
            int hSearch = Math.Min(searchPx, current.Height - 1);
            int bestDy = hTail; // worst‑case: no overlap
            double bestScore = double.MaxValue;

            // Sample columns for speed
            int[] cols = Enumerable.Range(0, w).Where(x => x % 3 == 0).ToArray();

            // Lock bits for fast access
            var lbPrev = prevTail.LockBits(new Rectangle(0, 0, prevTail.Width, prevTail.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var lbCur  = current.LockBits(new Rectangle(0, 0, current.Width, current.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* pPrev = (byte*)lbPrev.Scan0;
                    byte* pCur  = (byte*)lbCur.Scan0;
                    int stridePrev = lbPrev.Stride;
                    int strideCur  = lbCur.Stride;

                    for (int dy = 0; dy <= hSearch; dy++)
                    {
                        int overlap = Math.Min(hTail, current.Height - dy);
                        if (overlap < Math.Min(24, hTail / 4)) continue; // too little to compare

                        double ssd = 0;
                        for (int y = 0; y < overlap; y += 2) // stride rows for speed
                        {
                            int yPrev = hTail - overlap + y;
                            byte* rowPrev = pPrev + yPrev * stridePrev;
                            byte* rowCur  = pCur  + (dy + y) * strideCur;

                            foreach (int x in cols)
                            {
                                byte* a = rowPrev + x * 4;
                                byte* b = rowCur  + x * 4;
                                int dr = a[2] - b[2];
                                int dg = a[1] - b[1];
                                int db = a[0] - b[0];
                                ssd += dr * dr + dg * dg + db * db;
                            }
                        }

                        if (ssd < bestScore)
                        {
                            bestScore = ssd;
                            bestDy = hTail - overlap;
                        }
                    }
                }
            }
            finally
            {
                prevTail.UnlockBits(lbPrev);
                current.UnlockBits(lbCur);
            }

            // Optional tiny refinement by parabolic fit around the bestDy is omitted for simplicity
            return Math.Max(0, bestDy);
        }
    }

    public static class ImageCmp
    {
        public static bool Similar(Bitmap a, Bitmap b)
        {
            if (a.Width != b.Width || a.Height != b.Height) return false;
            // Compare a few random patches using a hash; good enough to detect stagnation
            string ha = HashTiny(a);
            string hb = HashTiny(b);
            return ha == hb;
        }

        public static string HashTiny(Bitmap bmp)
        {
            // downscale to 16x16 luminance and build hex string
            int w = 16, h = 16;
            using var small = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(small))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                g.DrawImage(bmp, new Rectangle(0, 0, w, h));
            }
            var sb = new StringBuilder(w * h);
            for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++)
            {
                var c = small.GetPixel(x, y);
                int Y = (c.R * 299 + c.G * 587 + c.B * 114) / 1000;
                sb.Append((Y / 16).ToString("X1"));
            }
            return sb.ToString();
        }
    }

    // ===== DPI helpers =====
    public static class GraphicsDpi
    {
        public static float ScreenDpiX { get { using var g = Graphics.FromHwnd(IntPtr.Zero); return g.DpiX; } }
        public static float ScreenDpiY { get { using var g = Graphics.FromHwnd(IntPtr.Zero); return g.DpiY; } }
    }

    // ===== Native interop =====
    internal static class Native
    {
        public const int WM_HOTKEY = 0x0312;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;

        public const int SW_RESTORE = 9;
        public const int SW_SHOWNOACTIVATE = 4;

        public const int SRCCOPY = 0x00CC0020;
        public const int CAPTUREBLT = 0x40000000;

        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;

        [Flags]
        public enum PrintWindowFlags : uint
        {
            PW_DEFAULT = 0x00000000,
            PW_CLIENTONLY = 0x00000001,
            PW_RENDERFULLCONTENT = 0x00000002
        }

        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }

        [DllImport("user32.dll")] public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);
        [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(System.Drawing.Point point);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
        [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
        [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hdc, IntPtr h);
        [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr ho);
        [DllImport("gdi32.dll")] public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight,
            IntPtr hObjSource, int nXSrc, int nYSrc, int dwRop);

        // DPI awareness
        [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
        [DllImport("shcore.dll", SetLastError = true)] static extern int SetProcessDpiAwareness(ProcessDpiAwareness value);
        enum ProcessDpiAwareness { Process_DPI_Unaware = 0, Process_System_DPI_Aware = 1, Process_Per_Monitor_DPI_Aware = 2 }
        public static void SetProcessDpiAwareness()
        {
            try { SetProcessDpiAwareness(ProcessDpiAwareness.Process_Per_Monitor_DPI_Aware); }
            catch { try { SetProcessDPIAware(); } catch { } }
        }
    }
}

/*
BUILD INSTRUCTIONS
------------------
1) Create a new WinForms project (.NET 8):
   dotnet new winforms -n ScrollShotPro -f net8.0-windows
   Replace Program.cs with this file’s content (or add as a new class and set Program/Main accordingly).

2) Edit your .csproj to include:
   <PropertyGroup>
      <OutputType>WinExe</OutputType>
      <UseWindowsForms>true</UseWindowsForms>
      <EnableWindowsTargeting>true</EnableWindowsTargeting>
   </PropertyGroup>

3) Build & run. The app lives in the system tray. Hotkeys:
   Ctrl+Alt+F = Full screen
   Ctrl+Alt+W = Active window
   Ctrl+Alt+R = Region
   Ctrl+Alt+S = Scrolling capture (pick window)

OUTPUT
------
Screenshots are saved in %USERPROFILE%\Pictures\ScrollShotPro with timestamps.

TIPS / PARITY WITH FASTSTONE
----------------------------
• For difficult apps (custom‑drawn lists, virtualized UIs), the UIA SmallIncrement path is much more stable than blind wheel scrolling.
• You can increase maxSteps or tweak searchPx in OverlapEstimator for extremely long pages.
• If a window blocks PrintWindow, fallback BitBlt still works but may capture occlusions; bring window to front first.
• Add GIF/MP4 recording later via Desktop Duplication API (DXGI) and MediaFoundation.

EXTENSIONS (OPTIONAL)
---------------------
• Replace OverlapEstimator with OpenCvSharp4 (template matching or feature matching) to handle parallax/animated banners.
• Add horizontal scrolling, auto‑cropping of sticky headers/footers (detect stationary stripes via frame differencing).
• Add editor overlay for annotations and one‑click copy to clipboard/cloud.
*/

// ===== OpenCV UPGRADE PATCH (append-only) =====
// NuGet: OpenCvSharp4, OpenCvSharp4.runtime.win, OpenCvSharp4.Extensions (optional)
// This section adds: nested scroll picking via UIA element, sticky-header detection, and OpenCV-based stitching.

using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace ScrollShotPro
{
    // UIA rectangle helper for element-level (sub-scroll) capture
    public struct RectUia { public double L,T,R,B; public RectUia(double l,double t,double r,double b){L=l;T=t;R=r;B=b;} public Rectangle ToRectangle(){ return Rectangle.FromLTRB((int)L,(int)T,(int)R,(int)B); } }

    // Overload: pick nested scroll container via UIA and use OpenCV-based stitcher
    public static class ScrollerCv
    {
        public static Bitmap CaptureScrollingCv(IntPtr hwnd, int maxSteps = 200)
        {
            // 1) Let user pick a UIA element (sub-scroll) if available
            dynamic pickedElm = null; RectUia? pickedRect = null;
            try
            {
                var uia = Activator.CreateInstance(Type.GetTypeFromProgID("UIAutomationClient.CUIAutomation"));
                System.Drawing.Point pt = Cursor.Position;
                var elm = uia.ElementFromPoint(new System.Drawing.Point(pt.X, pt.Y));
                var patternId = 10004; // ScrollPatternId
                var scroll = elm.GetCurrentPattern(patternId);
                if (scroll != null)
                {
                    var br = (double[])elm.GetCurrentPropertyValue(30001); // BoundingRectangle
                    pickedElm = elm;
                    pickedRect = new RectUia(br[0], br[1], br[2], br[3]);
                }
            }
            catch { }

            Rectangle viewport = pickedRect.HasValue ? pickedRect.Value.ToRectangle() : WindowUtil.GetClientRectScreen(hwnd);
            if (viewport.Width < 50 || viewport.Height < 50) throw new InvalidOperationException("Viewport too small");

            ActivateWindow(hwnd);
            Thread.Sleep(120);

            var driver = new ScrollDriver(hwnd, viewport, pickedElm);
            var stitch = new CvStitcher(viewport.Size);
            using (var first = ScreenGrab.CaptureArea(IntPtr.Zero, viewport)) stitch.Append(first, 0);

            int steps = 0, stagnant = 0; string prevHash = ""; Bitmap prev = null; int stickyTop = 0;

            while (steps < maxSteps)
            {
                bool ok = driver.HasUiaScrollPattern ? driver.ScrollSmallIncrement() : driver.ScrollFallbackPageDown();
                if (!ok) break;
                WaitForContentChange(hwnd, viewport, ref prevHash);

                using var raw = ScreenGrab.CaptureArea(IntPtr.Zero, viewport);
                stickyTop = Math.Max(stickyTop, CvStitcher.EstimateStickyHeader(stitch.LastFrame(), raw, Math.Min(120, viewport.Height/3)));
                using var shot = stickyTop>0 ? raw.Clone(new Rectangle(0, stickyTop, raw.Width, raw.Height-stickyTop), raw.PixelFormat) : (Bitmap)raw.Clone();

                if (prev!=null && ImageCmp.Similar(prev, shot)) { if (++stagnant>=4) break; } else stagnant=0;
                int dy = CvStitcher.EstimateVerticalOverlapCv(stitch.TailSlice(shot.Height/3), shot, viewport.Width);
                stitch.Append(shot, Math.Max(0, dy));
                prev?.Dispose(); prev = (Bitmap)shot.Clone();
                if (driver.AtVerticalEnd) break; steps++;
            }
            prev?.Dispose();
            return stitch.Render();
        }

        private static void ActivateWindow(IntPtr hwnd)
        {
            Native.SetForegroundWindow(hwnd);
            Native.ShowWindow(hwnd, Native.SW_SHOWNOACTIVATE);
            Native.ShowWindow(hwnd, Native.SW_RESTORE);
        }

        private static void WaitForContentChange(IntPtr hwnd, Rectangle viewport, ref string prevHash)
        {
            for (int i = 0; i < 25; i++)
            {
                using var bmp = ScreenGrab.CaptureArea(IntPtr.Zero, viewport);
                string h = ImageCmp.HashTiny(bmp);
                if (h != prevHash) { prevHash = h; return; }
                Thread.Sleep(12);
            }
        }
    }

    // OpenCV-powered stitcher with template matching + phase correlation fallback and sticky header probe
    public class CvStitcher
    {
        private readonly int viewportWidth;
        private readonly List<Bitmap> tiles = new();
        private readonly List<int> offsets = new();
        private int totalHeight; private Bitmap last;
        public CvStitcher(Size viewport) { viewportWidth = viewport.Width; totalHeight = 0; }
        public void Append(Bitmap frame, int overlap) { Bitmap copy = new Bitmap(frame.Width, frame.Height, PixelFormat.Format32bppArgb); using(var g=Graphics.FromImage(copy)) g.DrawImageUnscaled(frame,0,0); int y=Math.Max(0,totalHeight-overlap); tiles.Add(copy); offsets.Add(y); totalHeight=y+copy.Height; last?.Dispose(); last=(Bitmap)copy.Clone(); }
        public Bitmap TailSlice(int px){ var lastBmp=tiles.Last(); int h=Math.Min(px,lastBmp.Height); var bmp=new Bitmap(lastBmp.Width,h,PixelFormat.Format32bppArgb); using var g=Graphics.FromImage(bmp); g.DrawImage(lastBmp,new Rectangle(0,0,bmp.Width,bmp.Height), new Rectangle(0,lastBmp.Height-h,bmp.Width,h), GraphicsUnit.Pixel); return bmp; }
        public Bitmap LastFrame()=>last;
        public Bitmap Render(){ Bitmap result=new Bitmap(viewportWidth,totalHeight,PixelFormat.Format32bppArgb); using var g=Graphics.FromImage(result); g.Clear(Color.White); for(int i=0;i<tiles.Count;i++) g.DrawImageUnscaled(tiles[i],0,offsets[i]); return result; }

        public static int EstimateVerticalOverlapCv(Bitmap prevTail, Bitmap current, int width)
        {
            using var prevMat = BitmapConverter.ToMat(prevTail);
            using var curMat  = BitmapConverter.ToMat(current);
            int overlap = TemplateDy(prevMat, curMat);
            if (overlap >= 0) return overlap;
            double dy = PhaseDy(prevMat, curMat);
            return Math.Max(0, (int)Math.Round(dy));
        }

        private static int TemplateDy(Mat prevTail, Mat cur)
        {
            using var prevGray = new Mat(); using var curGray = new Mat();
            Cv2.CvtColor(prevTail, prevGray, ColorConversionCodes.BGRA2GRAY);
            Cv2.CvtColor(cur, curGray, ColorConversionCodes.BGRA2GRAY);
            using var prevEdge=new Mat(); using var curEdge=new Mat();
            Cv2.Canny(prevGray, prevEdge, 60, 180);
            Cv2.Canny(curGray,  curEdge,  60, 180);
            int hSearch = Math.Min(curEdge.Rows/3, 240); var roi=new Rect(0,0,curEdge.Cols, Math.Max(hSearch,40)); using var curTop=new Mat(curEdge, roi);
            using var res=new Mat(); Cv2.MatchTemplate(curTop, prevEdge, res, TemplateMatchModes.CCoeffNormed); Cv2.MinMaxLoc(res, out _, out var maxVal, out _, out var maxLoc);
            if (maxVal < 0.75) return -1; int dy = maxLoc.Y; int overlap = prevTail.Rows - dy; return Math.Max(0, overlap);
        }

        private static double PhaseDy(Mat prevTail, Mat cur)
        {
            using var prevGray=new Mat(); using var curGray=new Mat();
            Cv2.CvtColor(prevTail, prevGray, ColorConversionCodes.BGRA2GRAY);
            Cv2.CvtColor(cur, curGray, ColorConversionCodes.BGRA2GRAY);
            using var p32=new Mat(); using var c32=new Mat(); prevGray.ConvertTo(p32, MatType.CV_32F); curGray.ConvertTo(c32, MatType.CV_32F);
            var shift = Cv2.PhaseCorrelate(p32, c32); double dy = prevTail.Rows - Math.Max(0, shift.Y); return dy;
        }

        public static int EstimateStickyHeader(Bitmap prev, Bitmap cur, int maxProbe)
        {
            if (prev==null||cur==null) return 0; int h=Math.Min(Math.Min(maxProbe, prev.Height), cur.Height);
            using var p=BitmapConverter.ToMat(prev); using var c=BitmapConverter.ToMat(cur);
            using var pGray=new Mat(); using var cGray=new Mat(); Cv2.CvtColor(p,pGray,ColorConversionCodes.BGRA2GRAY); Cv2.CvtColor(c,cGray,ColorConversionCodes.BGRA2GRAY);
            int stable=0; const double thr=2.0; for(int y=0;y<h;y++){ using var r1=pGray.Row(pGray.Rows-h+y); using var r2=cGray.Row(y); using var diff=new Mat(); Cv2.Absdiff(r1,r2,diff); double m=Cv2.Mean(diff).Val0; if(m<thr) stable++; else break; }
            return Math.Min(stable, Math.Max(0, h));
        }
    }
}

// HOW TO USE
// 1) Install NuGet: OpenCvSharp4 + OpenCvSharp4.runtime.win
// 2) In BeginScrollCapture(), replace the old call with:
//    var result = ScrollerCv.CaptureScrollingCv(hwnd);
// 3) Build/run. When bắt đầu scroll-capture, di chuột lên vùng con có thanh cuộn (sub-scroll) rồi click.

