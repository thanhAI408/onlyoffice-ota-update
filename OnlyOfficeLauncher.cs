using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OnlyOfficeSuite
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    public class ClipboardBridgeForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll")]
        private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("user32.dll")]
        private static extern bool IsClipboardFormatAvailable(uint uFormat);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterClipboardFormat(string lpszFormat);

        [DllImport("gdi32.dll")]
        private static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[] lpbBuffer);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SetEnhMetaFileBits(uint cbBuffer, byte[] lpData);

        [DllImport("gdi32.dll")]
        private static extern bool PlayEnhMetaFile(IntPtr hdc, IntPtr hemf, ref RECT lpRect);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern UIntPtr GlobalSize(IntPtr hMem);

        private const int WM_CLIPBOARDUPDATE = 0x031D;
        private const uint CF_ENHMETAFILE = 14;
        private const uint GMEM_MOVEABLE = 0x0002;
        private const uint GMEM_ZEROINIT = 0x0040;

        private readonly uint _pngFormat;
        private readonly uint _htmlFormat;
        private readonly uint _markerFormat;
        private readonly uint _chemDrawFormat;
        private bool _isProcessing = false;
        private System.Windows.Forms.Timer _pollTimer;

        public ClipboardBridgeForm()
        {
            _pngFormat = RegisterClipboardFormat("PNG");
            _htmlFormat = RegisterClipboardFormat("HTML Format");
            _markerFormat = RegisterClipboardFormat("ONLYOFFICE_EMF_V6_HD");
            _chemDrawFormat = RegisterClipboardFormat("ChemDraw Interchange Format");

            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.Opacity = 0;
            this.Width = 1;
            this.Height = 1;

            TryConvertEmfClipboard();

            _pollTimer = new System.Windows.Forms.Timer();
            _pollTimer.Interval = 350;
            _pollTimer.Tick += (s, e) => TryConvertEmfClipboard();
            _pollTimer.Start();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
            if (!this.IsHandleCreated)
            {
                CreateHandle();
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            AddClipboardFormatListener(this.Handle);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            RemoveClipboardFormatListener(this.Handle);
            base.OnFormClosing(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                TryConvertEmfClipboard();
            }
            base.WndProc(ref m);
        }

        private static IntPtr AllocGlobalBytes(byte[] data)
        {
            IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, (UIntPtr)data.Length);
            if (hGlobal == IntPtr.Zero) return IntPtr.Zero;
            IntPtr ptr = GlobalLock(hGlobal);
            if (ptr == IntPtr.Zero) return IntPtr.Zero;
            try
            {
                Marshal.Copy(data, 0, ptr, data.Length);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }
            return hGlobal;
        }

        private static byte[] BuildCfHtml(string htmlFragment)
        {
            string html = "<html>\r\n<body>\r\n<!--StartFragment-->" + htmlFragment + "<!--EndFragment-->\r\n</body>\r\n</html>";
            string headerTpl = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
            int headerLen = string.Format(headerTpl, 0, 0, 0, 0).Length;
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            int startHtml = headerLen;
            int endHtml = headerLen + htmlBytes.Length;
            int startFrag = headerLen + Encoding.UTF8.GetByteCount("<html>\r\n<body>\r\n<!--StartFragment-->");
            int endFrag = startFrag + Encoding.UTF8.GetByteCount(htmlFragment);
            byte[] headerBytes = Encoding.UTF8.GetBytes(string.Format(headerTpl, startHtml, endHtml, startFrag, endFrag));
            byte[] result = new byte[headerBytes.Length + htmlBytes.Length + 1];
            Buffer.BlockCopy(headerBytes, 0, result, 0, headerBytes.Length);
            Buffer.BlockCopy(htmlBytes, 0, result, headerBytes.Length, htmlBytes.Length);
            result[result.Length - 1] = 0;
            return result;
        }

        private static bool RenderEmfBytesToBitmap(byte[] emfData, Bitmap bmp, int pad)
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                IntPtr hdc = g.GetHdc();
                try
                {
                    RECT rc = new RECT
                    {
                        Left = pad,
                        Top = pad,
                        Right = bmp.Width - pad,
                        Bottom = bmp.Height - pad
                    };
                    IntPtr hPlayEmf = SetEnhMetaFileBits((uint)emfData.Length, emfData);
                    if (hPlayEmf != IntPtr.Zero)
                    {
                        try
                        {
                            PlayEnhMetaFile(hdc, hPlayEmf, ref rc);
                        }
                        finally
                        {
                            DeleteEnhMetaFile(hPlayEmf);
                        }
                    }
                }
                finally
                {
                    g.ReleaseHdc(hdc);
                }
            }

            BitmapData bd = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            try
            {
                int stride = bd.Stride;
                int len = Math.Abs(stride) * bmp.Height;
                byte[] rgb = new byte[len];
                Marshal.Copy(bd.Scan0, rgb, 0, len);
                for (int y = 0; y < bmp.Height; y++)
                {
                    int row = y * stride;
                    for (int x = 0; x < bmp.Width * 3; x++)
                    {
                        if (rgb[row + x] < 250) return true;
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(bd);
            }
            return false;
        }

        private static void EnhanceContrastAndCrispness(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            BitmapData bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            try
            {
                int stride = bd.Stride;
                int len = Math.Abs(stride) * h;
                byte[] src = new byte[len];
                byte[] dst = new byte[len];
                Marshal.Copy(bd.Scan0, src, 0, len);
                Buffer.BlockCopy(src, 0, dst, 0, len);

                // Soft 4-neighbor stroke bolding + high-contrast darkening so 3x (288 DPI) chemical bonds and labels are bold, jet-black, and razor-sharp
                for (int y = 1; y < h - 1; y++)
                {
                    int row = y * stride;
                    int rowUp = (y - 1) * stride;
                    int rowDn = (y + 1) * stride;
                    for (int x = 1; x < w - 1; x++)
                    {
                        int p = row + x * 3;
                        int c = src[p];
                        int nMin = src[rowUp + x * 3];
                        if (src[rowDn + x * 3] < nMin) nMin = src[rowDn + x * 3];
                        if (src[p - 3] < nMin) nMin = src[p - 3];
                        if (src[p + 3] < nMin) nMin = src[p + 3];

                        int val = Math.Min(c, (c + nMin * 2) / 3);
                        if (val >= 242) val = 255;
                        else if (val <= 115) val = 0;
                        else val = (val - 115) * 255 / (242 - 115);

                        byte bVal = (byte)val;
                        dst[p] = bVal;
                        dst[p + 1] = bVal;
                        dst[p + 2] = bVal;
                    }
                }
                Marshal.Copy(dst, 0, bd.Scan0, len);
            }
            finally
            {
                bmp.UnlockBits(bd);
            }
        }

        public void TryConvertEmfClipboard()
        {
            if (_isProcessing) return;
            try
            {
                _isProcessing = true;

                if (!IsClipboardFormatAvailable(CF_ENHMETAFILE)) return;
                if (IsClipboardFormatAvailable(_markerFormat)) return;

                bool isChemDraw = IsClipboardFormatAvailable(_chemDrawFormat);
                if (!isChemDraw && IsClipboardFormatAvailable(_pngFormat) && IsClipboardFormatAvailable(_htmlFormat))
                {
                    return;
                }

                byte[] emfBytes = null;
                byte[] cdxBytes = null;

                bool opened = false;
                for (int retry = 0; retry < 10; retry++)
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        opened = true;
                        break;
                    }
                    Thread.Sleep(15);
                }
                if (!opened) return;

                try
                {
                    IntPtr hEmf = GetClipboardData(CF_ENHMETAFILE);
                    if (hEmf == IntPtr.Zero) return;

                    uint size = GetEnhMetaFileBits(hEmf, 0, null);
                    if (size < 88) return;
                    emfBytes = new byte[size];
                    if (GetEnhMetaFileBits(hEmf, size, emfBytes) == 0) return;

                    if (_chemDrawFormat != 0 && IsClipboardFormatAvailable(_chemDrawFormat))
                    {
                        IntPtr hCdx = GetClipboardData(_chemDrawFormat);
                        if (hCdx != IntPtr.Zero)
                        {
                            ulong sz = (ulong)GlobalSize(hCdx);
                            if (sz > 0 && sz < 8 * 1024 * 1024)
                            {
                                IntPtr ptr = GlobalLock(hCdx);
                                if (ptr != IntPtr.Zero)
                                {
                                    try
                                    {
                                        cdxBytes = new byte[(int)sz];
                                        Marshal.Copy(ptr, cdxBytes, 0, (int)sz);
                                    }
                                    finally
                                    {
                                        GlobalUnlock(hCdx);
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    CloseClipboard();
                }

                if (emfBytes == null || emfBytes.Length < 88) return;

                int boundsLeft = BitConverter.ToInt32(emfBytes, 8);
                int boundsTop = BitConverter.ToInt32(emfBytes, 12);
                int boundsRight = BitConverter.ToInt32(emfBytes, 16);
                int boundsBottom = BitConverter.ToInt32(emfBytes, 20);

                int frameLeft = BitConverter.ToInt32(emfBytes, 24);
                int frameTop = BitConverter.ToInt32(emfBytes, 28);
                int frameRight = BitConverter.ToInt32(emfBytes, 32);
                int frameBottom = BitConverter.ToInt32(emfBytes, 36);

                int devX = BitConverter.ToInt32(emfBytes, 72);
                int devY = BitConverter.ToInt32(emfBytes, 76);
                int mmX = BitConverter.ToInt32(emfBytes, 80);
                int mmY = BitConverter.ToInt32(emfBytes, 84);
                if (devX <= 0) devX = 1024;
                if (devY <= 0) devY = 768;
                if (mmX <= 0) mmX = 270;
                if (mmY <= 0) mmY = 203;

                int boundsW = boundsRight - boundsLeft;
                int boundsH = boundsBottom - boundsTop;

                int dispW = boundsW > 8 ? boundsW : (int)Math.Round(((frameRight - frameLeft) / 100.0) * 96.0 / 25.4);
                int dispH = boundsH > 8 ? boundsH : (int)Math.Round(((frameBottom - frameTop) / 100.0) * 96.0 / 25.4);

                if (dispW < 16) dispW = 240;
                if (dispH < 16) dispH = 160;
                if (dispW > 1200)
                {
                    double ratio = 1200.0 / dispW;
                    dispW = 1200;
                    dispH = Math.Max(32, (int)Math.Round(dispH * ratio));
                }

                int pad = 6;
                int totalDispW = dispW + pad * 2;
                int totalDispH = dispH + pad * 2;

                // Patch rclFrame (bytes 24..39) to match rclBounds (bytes 8..23) in 0.01mm units!
                byte[] patchedEmf = (byte[])emfBytes.Clone();
                if (boundsW > 0 && boundsH > 0)
                {
                    int newFL = (int)Math.Round(boundsLeft * 100.0 * mmX / devX);
                    int newFT = (int)Math.Round(boundsTop * 100.0 * mmY / devY);
                    int newFR = (int)Math.Round(boundsRight * 100.0 * mmX / devX);
                    int newFB = (int)Math.Round(boundsBottom * 100.0 * mmY / devY);
                    Buffer.BlockCopy(BitConverter.GetBytes(newFL), 0, patchedEmf, 24, 4);
                    Buffer.BlockCopy(BitConverter.GetBytes(newFT), 0, patchedEmf, 28, 4);
                    Buffer.BlockCopy(BitConverter.GetBytes(newFR), 0, patchedEmf, 32, 4);
                    Buffer.BlockCopy(BitConverter.GetBytes(newFB), 0, patchedEmf, 36, 4);
                }

                // Super-Resolution: Render at 6x native GDI vector resolution, then downsample with HighQualityBicubic to 3x (288 DPI)!
                int finalScale = 3;
                int superScale = 6;
                if (totalDispW * superScale > 4800 || totalDispH * superScale > 4800)
                {
                    finalScale = 2;
                    superScale = 4;
                }

                int superW = totalDispW * superScale;
                int superH = totalDispH * superScale;
                int finalW = totalDispW * finalScale;
                int finalH = totalDispH * finalScale;

                using (Bitmap superBmp = new Bitmap(superW, superH, PixelFormat.Format24bppRgb))
                using (Bitmap finalBmp = new Bitmap(finalW, finalH, PixelFormat.Format24bppRgb))
                {
                    bool drawn = RenderEmfBytesToBitmap(patchedEmf, superBmp, pad * superScale);
                    if (!drawn)
                    {
                        RenderEmfBytesToBitmap(emfBytes, superBmp, pad * superScale);
                    }

                    finalBmp.SetResolution(96f * finalScale, 96f * finalScale);
                    using (Graphics fg = Graphics.FromImage(finalBmp))
                    {
                        fg.Clear(Color.White);
                        fg.SmoothingMode = SmoothingMode.HighQuality;
                        fg.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        fg.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        fg.DrawImage(superBmp, new Rectangle(0, 0, finalW, finalH), new Rectangle(0, 0, superW, superH), GraphicsUnit.Pixel);
                    }

                    // Enhance contrast so chemical bonds and atom text are jet-black and ultra-sharp
                    EnhanceContrastAndCrispness(finalBmp);

                    byte[] pngBytes;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        finalBmp.Save(ms, ImageFormat.Png);
                        pngBytes = ms.ToArray();
                    }

                    // Build CF_HTML with 4x Ultra-HD PNG data and exact 1x display dimensions (totalDispW x totalDispH)
                    string b64 = Convert.ToBase64String(pngBytes);
                    string imgHtml = string.Format(
                        "<img src=\"data:image/png;base64,{0}\" width=\"{1}\" height=\"{2}\" style=\"width:{1}px;height:{2}px\" />",
                        b64, totalDispW, totalDispH);
                    byte[] cfHtmlBytes = BuildCfHtml(imgHtml);

                    // Step 1: Set standard Windows Bitmap/DIB via OLE Clipboard
                    Clipboard.SetImage(finalBmp);

                    // Step 2: Append HTML Format (so ONLYOFFICE uses 4x HD PNG at 1x display size!), raw PNG, original EMF, ChemDraw CDX, and V6 marker
                    for (int retry = 0; retry < 10; retry++)
                    {
                        if (OpenClipboard(IntPtr.Zero))
                        {
                            try
                            {
                                IntPtr hHtml = AllocGlobalBytes(cfHtmlBytes);
                                if (hHtml != IntPtr.Zero) SetClipboardData(_htmlFormat, hHtml);

                                IntPtr hPng = AllocGlobalBytes(pngBytes);
                                if (hPng != IntPtr.Zero) SetClipboardData(_pngFormat, hPng);

                                IntPtr hRestoredEmf = SetEnhMetaFileBits((uint)emfBytes.Length, emfBytes);
                                if (hRestoredEmf != IntPtr.Zero) SetClipboardData(CF_ENHMETAFILE, hRestoredEmf);

                                if (cdxBytes != null && cdxBytes.Length > 0)
                                {
                                    IntPtr hCdx = AllocGlobalBytes(cdxBytes);
                                    if (hCdx != IntPtr.Zero) SetClipboardData(_chemDrawFormat, hCdx);
                                }

                                IntPtr hMarker = AllocGlobalBytes(new byte[] { 1, 0 });
                                if (hMarker != IntPtr.Zero) SetClipboardData(_markerFormat, hMarker);
                            }
                            finally
                            {
                                CloseClipboard();
                            }
                            break;
                        }
                        Thread.Sleep(15);
                    }
                }
            }
            catch { }
            finally
            {
                _isProcessing = false;
            }
        }
    }

    public static class Program
    {
        private static string FindDesktopEditorsDir(string launcherDir)
        {
            string portableAppDir = Path.Combine(launcherDir, "App\\ONLYOFFICE\\DesktopEditors");
            string sysAppDir = "C:\\Program Files\\ONLYOFFICE\\DesktopEditors";
            string localAppDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs\\ONLYOFFICE\\DesktopEditors");

            if (File.Exists(Path.Combine(portableAppDir, "DesktopEditors.exe"))) return portableAppDir;
            if (File.Exists(Path.Combine(sysAppDir, "DesktopEditors.exe"))) return sysAppDir;
            if (File.Exists(Path.Combine(localAppDir, "DesktopEditors.exe"))) return localAppDir;
            return null;
        }

        private static void CreateShortcut(string lnkPath, string targetPath, string arguments, string workingDir, string iconPath, string description)
        {
            try
            {
                Type wshType = Type.GetTypeFromProgID("WScript.Shell");
                if (wshType == null) return;
                object wsh = Activator.CreateInstance(wshType);
                object sc = wshType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, wsh, new object[] { lnkPath });
                Type scType = sc.GetType();
                scType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, sc, new object[] { targetPath });
                scType.InvokeMember("Arguments", BindingFlags.SetProperty, null, sc, new object[] { arguments });
                scType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, sc, new object[] { workingDir });
                if (!string.IsNullOrEmpty(iconPath) && File.Exists(iconPath))
                {
                    scType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, sc, new object[] { iconPath + ",0" });
                }
                scType.InvokeMember("Description", BindingFlags.SetProperty, null, sc, new object[] { description });
                scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, sc, null);
            }
            catch { }
        }

        private static void RunInstall(string suiteRoot, string launcherDir, string selfExe)
        {
            try
            {
                string srcIcons = Path.Combine(suiteRoot, "icons");
                string dstIcons = Path.Combine(launcherDir, "App\\ONLYOFFICE\\DesktopEditors\\icons");
                Directory.CreateDirectory(dstIcons);
                if (Directory.Exists(srcIcons))
                {
                    foreach (string f in Directory.GetFiles(srcIcons, "*.ico"))
                    {
                        File.Copy(f, Path.Combine(dstIcons, Path.GetFileName(f)), true);
                    }
                }

                string appDir = FindDesktopEditorsDir(launcherDir);
                if (appDir == null)
                {
                    string msiPath = Path.Combine(Path.GetTempPath(), "ONLYOFFICE-DesktopEditors-x64.msi");
                    string url = "https://github.com/ONLYOFFICE/DesktopEditors/releases/latest/download/ONLYOFFICE-DesktopEditors-x64.msi";
                    ProcessStartInfo curlPsi = new ProcessStartInfo
                    {
                        FileName = "curl.exe",
                        Arguments = "-L --retry 3 -o \"" + msiPath + "\" \"" + url + "\"",
                        UseShellExecute = false
                    };
                    using (Process p = Process.Start(curlPsi))
                    {
                        if (p != null) p.WaitForExit();
                    }
                    if (File.Exists(msiPath))
                    {
                        ProcessStartInfo msiPsi = new ProcessStartInfo
                        {
                            FileName = "msiexec.exe",
                            Arguments = "/i \"" + msiPath + "\" /qn /norestart",
                            Verb = "runas",
                            UseShellExecute = true
                        };
                        using (Process p = Process.Start(msiPsi))
                        {
                            if (p != null) p.WaitForExit();
                        }
                        try { File.Delete(msiPath); } catch { }
                    }
                    appDir = FindDesktopEditorsDir(launcherDir);
                }

                try
                {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\ONLYOFFICE\\DesktopEditors"))
                    {
                        if (key != null)
                        {
                            key.SetValue("autoUpdateMode", "disable", RegistryValueKind.String);
                            key.SetValue("locale", "vi-VN", RegistryValueKind.String);
                            key.SetValue("lang", "vi-VN", RegistryValueKind.String);
                            key.SetValue("ignoreAssocMsg", "true", RegistryValueKind.String);
                        }
                    }
                }
                catch { }

                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string workDir = appDir ?? launcherDir;
                CreateShortcut(Path.Combine(desktop, "Word.lnk"), selfExe, "--create=word --keeplang=vi-VN", workDir, Path.Combine(dstIcons, "Word.ico"), "Word (ONLYOFFICE)");
                CreateShortcut(Path.Combine(desktop, "Excel.lnk"), selfExe, "--create=cell --keeplang=vi-VN", workDir, Path.Combine(dstIcons, "Excel.ico"), "Excel (ONLYOFFICE)");
                CreateShortcut(Path.Combine(desktop, "PowerPoint.lnk"), selfExe, "--create=slide --keeplang=vi-VN", workDir, Path.Combine(dstIcons, "PowerPoint.ico"), "PowerPoint (ONLYOFFICE)");
                CreateShortcut(Path.Combine(desktop, "PDF.lnk"), selfExe, "--keeplang=vi-VN", workDir, Path.Combine(dstIcons, "PDF.ico"), "PDF (ONLYOFFICE)");
                CreateShortcut(Path.Combine(desktop, "ONLYOFFICE Suite.lnk"), selfExe, "--keeplang=vi-VN", workDir, Path.Combine(dstIcons, "OfficeSuite.ico"), "ONLYOFFICE Suite");

                using (RegistryKey runKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    if (runKey != null && File.Exists(selfExe))
                    {
                        runKey.SetValue("OnlyOfficeChemDrawBridge", "\"" + selfExe + "\" --background", RegistryValueKind.String);
                    }
                }

                int myPid = Process.GetCurrentProcess().Id;
                foreach (Process proc in Process.GetProcessesByName("OnlyOfficeLauncher"))
                {
                    if (proc.Id != myPid)
                    {
                        try { proc.Kill(); proc.WaitForExit(1000); } catch { }
                    }
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = selfExe,
                    Arguments = "--background",
                    WorkingDirectory = launcherDir,
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
            }
            catch { }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            string launcherDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            string suiteRoot = Directory.GetParent(launcherDir) != null ? Directory.GetParent(launcherDir).FullName : launcherDir;
            string selfExe = Path.Combine(launcherDir, "OnlyOfficeLauncher.exe");

            bool backgroundOnly = false;
            bool installMode = false;
            List<string> editorArgs = new List<string>();
            if (args != null)
            {
                foreach (string a in args)
                {
                    if (string.Equals(a, "--background", StringComparison.OrdinalIgnoreCase))
                    {
                        backgroundOnly = true;
                    }
                    else if (string.Equals(a, "--install", StringComparison.OrdinalIgnoreCase))
                    {
                        installMode = true;
                    }
                    else
                    {
                        editorArgs.Add(a);
                    }
                }
            }

            if (installMode)
            {
                RunInstall(suiteRoot, launcherDir, selfExe);
                return;
            }

            try
            {
                using (RegistryKey runKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    if (runKey != null && File.Exists(selfExe))
                    {
                        runKey.SetValue("OnlyOfficeChemDrawBridge", "\"" + selfExe + "\" --background", RegistryValueKind.String);
                    }
                }
            }
            catch { }

            if (!backgroundOnly)
            {
                string portableAppDir = Path.Combine(launcherDir, "App\\ONLYOFFICE\\DesktopEditors");
                string appDir = FindDesktopEditorsDir(launcherDir);

                if (appDir != null)
                {
                    try
                    {
                        using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\ONLYOFFICE\\DesktopEditors"))
                        {
                            if (key != null)
                            {
                                key.SetValue("autoUpdateMode", "disable", RegistryValueKind.String);
                                key.SetValue("locale", "vi-VN", RegistryValueKind.String);
                                key.SetValue("lang", "vi-VN", RegistryValueKind.String);
                                key.SetValue("ignoreAssocMsg", "true", RegistryValueKind.String);
                            }
                        }
                    }
                    catch { }

                    if (string.Equals(appDir, portableAppDir, StringComparison.OrdinalIgnoreCase))
                    {
                        string portableProfile = Path.Combine(launcherDir, "PortableData");
                        Directory.CreateDirectory(Path.Combine(portableProfile, "AppData\\Local\\ONLYOFFICE\\DesktopEditors\\data"));
                        Directory.CreateDirectory(Path.Combine(portableProfile, "AppData\\Roaming"));
                        Directory.CreateDirectory(Path.Combine(portableProfile, "Documents"));
                        Directory.CreateDirectory(Path.Combine(portableProfile, "Desktop"));

                        Environment.SetEnvironmentVariable("USERPROFILE", portableProfile);
                        Environment.SetEnvironmentVariable("LOCALAPPDATA", Path.Combine(portableProfile, "AppData\\Local"));
                        Environment.SetEnvironmentVariable("APPDATA", Path.Combine(portableProfile, "AppData\\Roaming"));
                    }

                    StringBuilder argBuilder = new StringBuilder();
                    foreach (string a in editorArgs)
                    {
                        if (argBuilder.Length > 0) argBuilder.Append(" ");
                        argBuilder.Append(a.Contains(" ") ? "\"" + a + "\"" : a);
                    }
                    if (argBuilder.Length == 0) argBuilder.Append("--keeplang=vi-VN");

                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = Path.Combine(appDir, "DesktopEditors.exe"),
                        Arguments = argBuilder.ToString(),
                        WorkingDirectory = appDir,
                        UseShellExecute = false
                    };
                    try
                    {
                        Process.Start(psi);
                    }
                    catch { }
                }
            }

            try
            {
                string oldExe = Path.Combine(launcherDir, "OnlyOfficeLauncher.old.exe");
                if (File.Exists(oldExe)) File.Delete(oldExe);
            }
            catch { }

            ThreadPool.QueueUserWorkItem(_ => CheckRemoteUpdate(suiteRoot, launcherDir));

            bool createdNew;
            using (Mutex mutex = new Mutex(true, "Local\\OnlyOfficeChemDrawClipboardBridge", out createdNew))
            {
                if (createdNew)
                {
                    Application.Run(new ClipboardBridgeForm());
                }
            }
        }

        private static void CopyDirectoryRecursive(string sourceDir, string targetDir)
        {
            Directory.CreateDirectory(targetDir);
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                try
                {
                    string destFile = Path.Combine(targetDir, Path.GetFileName(file));
                    File.Copy(file, destFile, true);
                }
                catch { }
            }
            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                CopyDirectoryRecursive(subDir, Path.Combine(targetDir, Path.GetFileName(subDir)));
            }
        }

        private static void CheckRemoteUpdate(string suiteRoot, string launcherDir)
        {
            try
            {
                string urlFile = Path.Combine(suiteRoot, "link-cap-nhat-tu-dong.txt");
                if (!File.Exists(urlFile)) urlFile = Path.Combine(launcherDir, "link-cap-nhat-tu-dong.txt");
                string rawUrl = "https://raw.githubusercontent.com/thanhAI408/onlyoffice-ota-update/main/Gui-Cho-May-Khac-Sieu-Nhe.zip";
                if (File.Exists(urlFile))
                {
                    foreach (string line in File.ReadAllLines(urlFile))
                    {
                        string trimmed = line.Trim();
                        if (!string.IsNullOrEmpty(trimmed) && !trimmed.StartsWith("#"))
                        {
                            rawUrl = trimmed;
                            break;
                        }
                    }
                }
                if (string.IsNullOrEmpty(rawUrl) || !rawUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return;

                string downloadUrl = rawUrl;
                Match m = Regex.Match(rawUrl, @"/file/d/([a-zA-Z0-9_-]+)");
                if (m.Success)
                {
                    downloadUrl = "https://drive.google.com/uc?export=download&id=" + m.Groups[1].Value;
                }
                else
                {
                    downloadUrl += (downloadUrl.Contains("?") ? "&t=" : "?t=") + DateTime.UtcNow.Ticks.ToString();
                }

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                string tmpZip = Path.Combine(Path.GetTempPath(), "OO_AutoUpdate_Patch.zip");
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");
                    wc.Headers.Add("Cache-Control", "no-cache");
                    wc.DownloadFile(downloadUrl, tmpZip);
                }

                FileInfo fi = new FileInfo(tmpZip);
                if (!fi.Exists || fi.Length < 1024) return;

                byte[] magic = new byte[4];
                using (FileStream fs = File.OpenRead(tmpZip))
                {
                    if (fs.Read(magic, 0, 4) < 4 || magic[0] != 0x50 || magic[1] != 0x4B) return;
                }

                string newHash;
                using (SHA256 sha = SHA256.Create())
                using (FileStream fs = File.OpenRead(tmpZip))
                {
                    newHash = BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "");
                }

                string stateFile = Path.Combine(launcherDir, "last-update.sha256");
                string oldHash = File.Exists(stateFile) ? File.ReadAllText(stateFile).Trim() : "";
                if (string.Equals(newHash, oldHash, StringComparison.OrdinalIgnoreCase))
                {
                    try { File.Delete(tmpZip); } catch { }
                    return;
                }

                File.WriteAllText(stateFile, newHash);

                string tmpExtract = Path.Combine(Path.GetTempPath(), "OO_Patch_Ext_" + Guid.NewGuid().ToString("N").Substring(0, 6));
                ZipFile.ExtractToDirectory(tmpZip, tmpExtract);

                string srcRoot = tmpExtract;
                string[] subDirs = Directory.GetDirectories(tmpExtract);
                if (subDirs.Length == 1 && Path.GetFileName(subDirs[0]).IndexOf("ONLYOFFICE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    srcRoot = subDirs[0];
                }

                CopyDirectoryRecursive(srcRoot, suiteRoot);

                string csc = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe";
                if (!File.Exists(csc)) csc = "C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319\\csc.exe";
                string csFile = Path.Combine(launcherDir, "OnlyOfficeLauncher.cs");
                string exeFile = Path.Combine(launcherDir, "OnlyOfficeLauncher.exe");
                string oldExe = Path.Combine(launcherDir, "OnlyOfficeLauncher.old.exe");

                if (File.Exists(csc) && File.Exists(csFile))
                {
                    try
                    {
                        if (File.Exists(oldExe)) File.Delete(oldExe);
                        if (File.Exists(exeFile)) File.Move(exeFile, oldExe);
                    }
                    catch { }

                    ProcessStartInfo cscPsi = new ProcessStartInfo
                    {
                        FileName = csc,
                        Arguments = "/nologo /target:winexe /optimize+ /out:\"" + exeFile + "\" /r:System.dll,System.Drawing.dll,System.Windows.Forms.dll,System.IO.Compression.FileSystem.dll \"" + csFile + "\"",
                        CreateNoWindow = true,
                        UseShellExecute = false
                    };
                    using (Process p = Process.Start(cscPsi))
                    {
                        if (p != null) p.WaitForExit(15000);
                    }
                }

                try { Directory.Delete(tmpExtract, true); } catch { }
                try { File.Delete(tmpZip); } catch { }
            }
            catch { }
        }
    }
}
