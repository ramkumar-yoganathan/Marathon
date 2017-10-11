using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Marathon
{
    public class CrashHandler
    {
        private const uint WM_CLOSE = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static Logger Log = Logger.GetLogger(typeof(CrashHandler));

        public static void HandleCrashWindows()
        {
            try
            {
                string crashSnapshotName = string.Empty;
                IList<string> crashWindows = TestConfiguration.CrashProcess.Split(TestConfiguration.LibrarySplitChar).ToList();
                foreach (KeyValuePair<IntPtr, string> window in OpenWindowGetter.GetOpenWindows())
                {
                    string windowTitle = window.Value;
                    foreach (string crashWindow in crashWindows)
                    {
                        if (windowTitle.Trim().ToUpper() == crashWindow.Trim().ToUpper())
                        {
                            IntPtr windowHandle = window.Key;
                            SetForegroundWindow(windowHandle);
                            crashSnapshotName = Logger.CaptureScreenshot();
                            PostMessage(windowHandle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                            TestConfiguration.CrashListTable.Add(crashSnapshotName);
                        }
                    }
                }
            }
            catch (Exception)
            {
                //Ignore the exception. We are not considering at the moment.
            }
        }
    }
}