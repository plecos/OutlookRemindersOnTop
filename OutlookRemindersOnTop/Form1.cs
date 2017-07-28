using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookRemindersOnTop
{
    public partial class Form1 : Form
    {
        #region User32 API
        [DllImport("user32.dll")]
        static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDesktopWindowsDelegate lpfn, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "GetWindowText", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumDesktopWindowsDelegate enumProc, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        internal enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }
        #endregion


        private delegate bool EnumDesktopWindowsDelegate(IntPtr hWnd, int lParam);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const UInt32 SWP_NOSIZE = 0x0001;
        private const UInt32 SWP_NOMOVE = 0x0002;
        private const UInt32 SWP_SHOWWINDOW = 0x0040;
        private const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW;
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int SW_RESTORE = 9;

        public Form1()
        {
            InitializeComponent();

            timer1.Enabled = true;

            RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            string flag = (string)rk.GetValue("OutlookRemindersOnTop","");
            rk.SetValue("OutlookRemindersOnTop", flag);
            if (flag == "") startAtLoginToolStripMenuItem.Checked = false; else startAtLoginToolStripMenuItem.Checked = true;

            notifyIcon1.ShowBalloonTip(5000, "Outlook Reminders On Top", "Program is running...", ToolTipIcon.Info);
        }
    

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var collection = new List<IntPtr>();
            EnumDesktopWindowsDelegate filter = delegate (IntPtr hWnd, int lParam)
            {
                StringBuilder strbTitle = new StringBuilder(255);
                int nLength = GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                string strTitle = strbTitle.ToString();
                if (strTitle.Contains("0 Reminder"))
                    return true;
                if (strTitle.Contains("Reminder"))
                    collection.Add(hWnd);
                return true;
            };

            if (EnumWindows(filter, IntPtr.Zero))
            {
                foreach (IntPtr hWnd in collection)
                {
                    uint pid = 0;
                    GetWindowThreadProcessId(hWnd, out pid);
                    Process p = Process.GetProcessById((int)pid);
                    try
                    {
                        FileInfo fi = new FileInfo(p.MainModule.FileName);
                        if (fi.Name.ToLower() == "outlook.exe")
                        {
                            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
                            placement.length = Marshal.SizeOf(placement);
                            GetWindowPlacement(hWnd, ref placement);
                            if (placement.showCmd == ShowWindowCommands.Hide)
                            {
                                SetActiveWindow(hWnd);
                                ShowWindow(hWnd, SW_RESTORE);
                            }
                            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS);
                        }
                    }
                    catch (Win32Exception)
                    {

                    }
                    
                }
            }
        }

        private void startAtLoginToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            string flag = "";
            if (startAtLoginToolStripMenuItem.Checked)
                flag = "";
            else flag = Environment.CommandLine;
            rk.SetValue("OutlookRemindersOnTop", flag);
            if (flag == "") startAtLoginToolStripMenuItem.Checked = false; else startAtLoginToolStripMenuItem.Checked = true;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            MessageBox.Show( "Written by Ken Salter\n\nCopyright 2017\n\nVersion " + v.Major+"."+ v.Minor+"."+v.Build+"."+v.Revision, "Outlook Reminders On Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
