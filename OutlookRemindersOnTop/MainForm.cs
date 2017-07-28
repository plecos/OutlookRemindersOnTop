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
    public partial class MainForm : Form
    {
        #region Win32 API
        [DllImport("user32.dll", EntryPoint = "GetWindowText", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsDelegate enumProc, IntPtr lParam);
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
        [DllImport("kernel32.dll")]
        static extern IntPtr OpenProcess(ProcessAccessFlags dwDesiredAccess, bool bInheritHandle, uint dwProcessId);
        [DllImport("psapi.dll")]
        static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In] [MarshalAs(UnmanagedType.U4)] int nSize);
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseHandle(IntPtr hObject);

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

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }


        private delegate bool EnumWindowsDelegate(IntPtr hWnd, int lParam);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const UInt32 SWP_NOSIZE = 0x0001;
        private const UInt32 SWP_NOMOVE = 0x0002;
        private const UInt32 SWP_SHOWWINDOW = 0x0040;
        private const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW;
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int SW_RESTORE = 9;

        #endregion

        #region Constructor
        public MainForm()
        {
            InitializeComponent();

            // enable the timer that will check for reminder windows
            monitorWindowsTimer.Enabled = true;

            // get the current status of the Startup At Login and set the context menu item check
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            string flag = (string)rk.GetValue("OutlookRemindersOnTop","");
            rk.SetValue("OutlookRemindersOnTop", flag);
            if (flag == "") startAtLoginToolStripMenuItem.Checked = false; else startAtLoginToolStripMenuItem.Checked = true;

        }
        #endregion

        #region Overrides
        /// <summary>
        /// This override will allow us to hide the form, since we don't use it
        /// </summary>
        /// <param name="value">flag to indicate to show the form</param>
        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }
        #endregion

        #region Context Menu Item click events
        /// <summary>
        /// This function is called when the user selectes the Exit menu item
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();            
        }

        /// <summary>
        /// Toggle the Start at Login
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void startAtLoginToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // read the registry entry for this program to auto start at login
            // then flip the value to the opposite, write it back, and set the context menu item check
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            string flag = "";
            if (startAtLoginToolStripMenuItem.Checked)
                flag = "";
            else flag = Environment.CommandLine;
            rk.SetValue("OutlookRemindersOnTop", flag);
            if (flag == "") startAtLoginToolStripMenuItem.Checked = false; else startAtLoginToolStripMenuItem.Checked = true;
        }

        /// <summary>
        /// Display a simple About screen
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            MessageBox.Show("Written by Ken Salter\n\nCopyright 2017\n\nVersion " + v.Major + "." + v.Minor + "." + v.Build + "." + v.Revision, "Outlook Reminders On Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region Timer Events
        /// <summary>
        /// This function fires every 500ms.  It will scan all windows looking for the Outlook Reminder window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void monitorWindowsTimer_Tick(object sender, EventArgs e)
        {
            EnumWindowsDelegate filter = delegate (IntPtr hWnd, int lParam)
            {
                try
                {
                    // get the process id of the window we are looking at
                    uint pid = 0;
                    GetWindowThreadProcessId(hWnd, out pid);
                    StringBuilder fileName = new StringBuilder(255);
                    IntPtr hProcess = OpenProcess(ProcessAccessFlags.QueryInformation|ProcessAccessFlags.VirtualMemoryRead, false, pid);
                    uint ret = GetModuleFileNameEx(hProcess, IntPtr.Zero, fileName, fileName.Capacity);
                    if (ret == 0)
                    {
                        int lastError = Marshal.GetLastWin32Error();
                        Win32Exception ex = new Win32Exception(lastError);
                        Trace.WriteLine(ex.Message);
                        CloseHandle(hProcess);
                        return true;

                    }
                    CloseHandle(hProcess);

                    // we only care if this window belongs to Outlook
                    if (fileName.ToString().ToLower().Contains("outlook.exe"))
                    {
                        // retrieve the text of the Window
                        StringBuilder strbTitle = new StringBuilder(255);
                        int nLength = GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                        string strTitle = strbTitle.ToString();

                        // ignore the "0 Reminder" window
                        if (strTitle.Contains("0 Reminder"))
                            return true;
                        if (strTitle.Contains("Reminder"))
                            // we have a winner - process it
                            MoveWindowOnTop(hWnd);
                    }

                }
                catch (Exception ex)
                {
                    // for any exception, display it and kill the timer
                    monitorWindowsTimer.Enabled = false;
                    MessageBox.Show("Fatal error encountered!\n\n"+ex.Message + "\n\n" + ex.StackTrace, "Exception in Outlook Reminders On Top", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
               
                return true;
                
            };

            EnumWindows(filter, IntPtr.Zero);
        }

        /// <summary>
        /// This function will move the window to be on top, and unhide it if was hidden
        /// </summary>
        /// <param name="hWnd">Handle to the window</param>
        private void MoveWindowOnTop(IntPtr hWnd)
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
        #endregion

    }
}
