using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace Marathon
{
    internal class TestMonitor
    {
        private static Logger Log = Logger.GetLogger(typeof(TestMonitor));

        /// <summary>
        /// Check the UFT is runnig or not
        /// </summary>
        /// <returns></returns>
        public static bool IsUftProcessExists()
        {
            bool uftProcess = Process.GetProcesses().Any(p => p.ProcessName.Contains("UFT"));
            return uftProcess;
        }

        /// <summary>
        /// Check the UFT is runnig or not
        /// </summary>
        /// <returns></returns>
        public static bool IsUftVisible()
        {
            bool isUftWindowExists =
                Process.GetProcesses().Any(p => p.MainWindowTitle.Contains(TestConfiguration.UftWindowTitle));
            return isUftWindowExists;
        }

        public static bool IsUftNotRespondingState()
        {
            bool isUftNotResponding =
                Process.GetProcesses().Any(p => p.MainWindowTitle.Contains(TestConfiguration.UftTestNotRespondingTitle));
            return isUftNotResponding;
        }

        /// <summary>
        /// This is an Watch method to check the uft process perodically about the running status.
        /// </summary>
        public static void WatchExecution(string currentSessionFolder)
        {
            string defaultSpreadSheetPath = Path.Combine(currentSessionFolder, TestConfiguration.DefaultSpreadsheet);
            var uftRunSessionStartTime = DateTime.Now;
            int timeElapsed = (DateTime.Now - uftRunSessionStartTime).Minutes;
            try
            {
                while (timeElapsed < TestConfiguration.TestRunSessionTimeout)
                {
                    CrashHandler.HandleCrashWindows();
                    timeElapsed = (DateTime.Now - uftRunSessionStartTime).Minutes;
                    if (File.Exists(defaultSpreadSheetPath))
                    {
                        break;
                    }
                    else
                    {
                        //Enable this when you are  in the need of debugging.
                        //Log.Normal("Test is running since " + timeElapsed + " minutes.");
                    }
                    Thread.Sleep(60000);
                }
            }
            catch (Exception exception)
            {
                Log.Error(exception);
            }
            finally
            {
                if (timeElapsed > TestConfiguration.TestRunSessionTimeout)
                {
                    Log.Warning("Test is running more then expected. Stopping the test now.");
                    bool isUftRunning = UnifiedFunctionalTesting.UftTest.IsRunning;
                    if (isUftRunning)
                    {
                        string beforeCloseLongRunning = Logger.CaptureScreenshot();
                        TestConfiguration.CrashListTable.Add(beforeCloseLongRunning);
                        bool isResponding = Utilities.GetRespondingState("UFT");
                        if (!isResponding)
                        {
                            CrashHandler.HandleCrashWindows();
                            Utilities.TerminateDeadLockBrowsers();
                        }
                        UnifiedFunctionalTesting.UftTest.Stop();
                        UnifiedFunctionalTesting.UftTest.Close();
                    }
                }
            }
        }

        public static void ExecuteParallel(params Action[] tasks)
        {
            // Initialize the reset events to keep track of completed threads
            ManualResetEvent[] resetEvents = new ManualResetEvent[tasks.Length];

            // Launch each method in it's own thread
            for (int i = 0; i < tasks.Length; i++)
            {
                resetEvents[i] = new ManualResetEvent(false);
                ThreadPool.QueueUserWorkItem(index =>
                {
                    int taskIndex = (int)index;

                    // Execute the method
                    tasks[taskIndex]();

                    // Tell the calling thread that we're done
                    resetEvents[taskIndex].Set();
                }, i);
            }

            // Wait for all threads to execute
            WaitHandle.WaitAll(resetEvents);
        }
    }
}