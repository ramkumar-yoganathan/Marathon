using QuickTest;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Marathon
{
    public static class UnifiedFunctionalTesting
    {
        private static readonly Logger Log = Logger.GetLogger(typeof(UnifiedFunctionalTesting));

        [CLSCompliant(false)]
        public static Application UftApplication;

        [CLSCompliant(false)]
        public static Test UftTest;

        //private static object errorDescription;

        /// <summary>
        ///     Create an instance of Qt Test settings COM Object
        /// </summary>
        private static TestSettings UftTestSettings => UftTest.Settings;

        /// <summary>
        ///     Create an instance of Qt Test settings COM Object
        /// </summary>
        private static LocalSystemMonitorSettings UftLocalSystemMonitor => UftTest.Settings.LocalSystemMonitor;

        /// <summary>
        ///     Create an instance of Qt  options COM Object
        /// </summary>
        private static Options UftRunOptions => UftApplication.Options;

        /// <summary>
        ///     Create an instance of Qt  options COM Object
        /// </summary>
        private static AutoExportReportConfigOptions UftAutoExport => UftApplication.Options.Run.AutoExportReportConfig;

        /// <summary>
        ///     Create an instance of Qt  options COM Object
        /// </summary>
        private static TestLibraries UftTestLibraries => UftTest.Settings.Resources.Libraries;

        /// <summary>
        ///     Create an intance of Qt Folder Com Object
        /// </summary>
        private static FoldersOptions UftTestFolders => UftApplication.Folders;

        /// <summary>
        ///     Launch the Uft application.
        /// </summary>
        private static void Launch()
        {
            try
            {
                if (UftApplication == null)
                {
                    UftApplication = new Application();
                }
                if (!UftApplication.Launched)
                {
                    SetAddInsAssociated();
                    UftApplication.Launch();
                    UftApplication.Visible = TestConfiguration.ShowAutomationTool;
                }
            }
            catch (Exception exception)
            {
                Log.Failure(
                    "Automation tool heart beat reported fatal error.Quitting exeuction session.Please check the tool and run again.");
                Log.Error(exception);
                Console.WriteLine(exception.StackTrace);
                Environment.Exit(0);
            }
        }

        /// <summary>
        ///     Open the given test name.
        /// </summary>
        /// <param name="qtTestName">Name of the test</param>
        /// <returns>Test Object</returns>
        private static Test Open(string qtTestName)
        {
            List<string> qtEnvironment = new List<string>();
            var qtTestNameFullPath = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.UiScripts, qtTestName);
            try
            {
                if (Directory.Exists(qtTestNameFullPath))
                {
                    if (File.Exists(TestConfiguration.TestDataDrivenFile))
                    {
                        qtEnvironment.Add("RowNum" + "=" + TestConfiguration.RowCounter);
                        qtEnvironment.Add("env" + "=" + TestConfiguration.Environment);
                        PrepareEnvironmentInfo(qtEnvironment);
                        InjectEnvironmentInfo(Path.Combine(qtTestNameFullPath, "Action1", "Script.mts"), false);
                    }
                    UftApplication.Open(qtTestNameFullPath);
                }
                else
                {
                    Log.Warning("Test name is not found in the location", qtTestNameFullPath);
                }
            }
            catch (Exception exception)
            {
                Log.Error(exception);
            }
            return UftApplication.Test;
        }

        /// <summary>
        ///     Close UFT
        /// </summary>
        /// <param name="running"></param>
        /// <param name="close"></param>
        /// <param name="quitTool"></param>
        public static void Close(bool running = false, bool close = true, bool quitTool = true)
        {
            try
            {
                var isUftVisible = TestMonitor.IsUftVisible();
                if (isUftVisible)
                {
                    if (running)
                    {
                        UftTest.Stop();
                        Thread.Sleep(5000);
                    }

                    if (quitTool)
                    {
                        if (UftApplication == null)
                        {
                            UftApplication = new Application();
                            UftApplication.Quit();
                            Thread.Sleep(5000);
                            UftApplication = null;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Log.Error(
                    "Unable to close/stop the long running. Possible with dead lock state. Terminating UFT to clear deadlock.");
                Log.Error(exception);
            }
            finally
            {
                Release();
            }
        }

        /// <summary>
        ///     Set Test Setttings Values. This is same as File -> Test Settings -> Run (Treenode)
        /// </summary>
        /// <param name="settings"></param>
        private static void TestSettings(Settings settings)
        {
            switch (settings)
            {
                case Settings.Run:
                    UftTestSettings.Run.OnError = TestConfiguration.OnError;
                    UftTestSettings.Run.ObjectSyncTimeOut = TestConfiguration.ObjectSyncTimeout;
                    UftTestSettings.Run.IterationMode = TestConfiguration.IterationMode;
                    break;

                case Settings.LocalSystemMonitor:
                    UftLocalSystemMonitor.ApplicationName = TestConfiguration.MonitorApplicationProcess;
                    UftLocalSystemMonitor.Enable = true;
                    UftLocalSystemMonitor.SystemCounters.RemoveAll();
                    UftLocalSystemMonitor.SystemCounters.Add("Memory Usage (in MB)", -1);
                    UftLocalSystemMonitor.SystemCounters.Add("% Processor Time", -1);
                    break;

                case Settings.Environment:
                    UftTest.Environment("RowNum").Value = 2;
                    break;
            }
            UftTest.Save();
        }

        ///<summary>
        ///
        ///</summary>
        private static void SetAddInsAssociated()
        {
            int addinsCount = UftApplication.Addins.Count;
        }

        /// <summary>
        ///     Set run time options for Run, Screen Capture, AutoExport and Web Advacned. This settings will be
        ///     reflected in the ui Tools -> Options
        /// </summary>
        /// <param name="options">RunTimeOptions enum</param>
        private static void RunOptions(RuntimeOption options)
        {
            string customXslPath = Path.Combine(TestConfiguration.FrameworkWorkspace, TestConfiguration.ReportXslRrv);
            switch (options)
            {
                case RuntimeOption.RunMode:
                    UftRunOptions.Run.ViewResults = TestConfiguration.ViewResults;
                    UftRunOptions.Run.RunMode = TestConfiguration.RunMode;
                    UftRunOptions.Run.ReportFormat = TestConfiguration.ReportFormat;
                    break;

                case RuntimeOption.ScreenCapture:
                    UftRunOptions.Run.ImageCaptureForTestResults = TestConfiguration.ImageCaptureForTestResults;
                    UftRunOptions.Run.MovieCaptureForTestResults = TestConfiguration.MovieCaptureForTestResults;
                    UftRunOptions.Run.MovieSegmentSize = TestConfiguration.MovieSegmentSize;
                    UftRunOptions.Run.SaveMovieOfEntireRun = TestConfiguration.SaveMovieOfEntireRun;
                    UftRunOptions.Run.StepExecutionDelay = TestConfiguration.StepExecutionDelay;
                    break;

                case RuntimeOption.AutoExport:
                    UftAutoExport.AutoExportResults = true;
                    UftAutoExport.StepDetailsReport = true;
                    UftAutoExport.LogTrackingReport = true;
                    UftAutoExport.ScreenRecorderReport = false;
                    UftAutoExport.SystemMonitorReport = false;
                    UftAutoExport.ExportLocation = string.Empty;
                    UftAutoExport.UserDefinedXSL = customXslPath;
                    UftAutoExport.StepDetailsReportType = TestConfiguration.StepDetailsReportType;
                    //UftAutoExport.StepDetailsReportFormat = Configuration.StepDetailsReportFormat;
                    UftAutoExport.ExportForFailedRunsOnly = false;
                    break;

                case RuntimeOption.WebAdvanced:
                    UftRunOptions.Web.BrowserCleanup = true;
                    UftRunOptions.Web.RunOnlyClick = false;
                    UftRunOptions.Web.RunUsingSourceIndex = true;
                    UftRunOptions.Web.RunMouseByEvents = false;
                    break;

                case RuntimeOption.Folders:
                    UftApplication.Folders.RemoveAll();
                    break;
            }
            UftTest.Save();
        }

        /// <summary>
        ///     Associate functional libraries in to the test
        /// </summary>
        private static void AddResources()
        {
            var libraries = "";
            StringBuilder libraryList = new StringBuilder();
            if (TestConfiguration.OverWritedefault)
            {
                string businessLibraries = TestConfiguration.TestResourcesList.Split('|')[0];
                List<string> userlibraries = TestConfiguration.TestResourcesList.Split('|')[1].Split(';').ToList();
                foreach (string userLibrary in userlibraries)
                {
                    if (!string.IsNullOrEmpty(userLibrary))
                    {
                        string userLibraryValue = userLibrary.Replace("Replace", "").Replace("(", "").Replace(")", "");
                        libraryList.Append(userLibraryValue.Split(',')[1]).Append(TestConfiguration.LibrarySplitChar);
                        string commonLibraries = TestConfiguration.CommonLibsTestResourcesList.ToLower();
                        commonLibraries = commonLibraries.Replace(userLibraryValue.Split(',')[0].ToLower(), "");
                        TestConfiguration.CommonLibsTestResourcesList = commonLibraries;
                    }
                }
                libraries = libraryList
                    .Append(businessLibraries)
                    .Append(TestConfiguration.LibrarySplitChar)
                    .Append(TestConfiguration.CommonLibsTestResourcesList).ToString();
            }
            else
            {
                libraries = libraryList
                                    .Append(TestConfiguration.TestResourcesList)
                                    .Append(TestConfiguration.LibrarySplitChar)
                                    .Append(TestConfiguration.CommonLibsTestResourcesList).ToString();
            }
            var librariesList = libraries.Split(TestConfiguration.LibrarySplitChar).ToList();
            UftTestLibraries.RemoveAll();
            foreach (string library in librariesList)
            {
                UftTestLibraries.Add(library);
            }
            UftTest.Save();
        }

        private static void AddTestSourcesFolders()
        {
            string businessLibsFolder = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.LibraryFolder);
            UftTestFolders.RemoveAll();
            UftTestFolders.Add(TestConfiguration.FrameworkWorkspace);
            UftTestFolders.Add(businessLibsFolder);
            UftTest.Save();
        }

        /// <summary>
        ///     Run the test
        /// </summary>
        /// <param name="testName">Name of the test to be run</param>
        public static ResultStatus Run(string testName)
        {
            var qtTestNameFullPath = "";
            var lastRunResultStatus = ResultStatus.Inconclusive;
            try
            {
                ReportFormats reportFormat;
                Enum.TryParse(TestConfiguration.ReportFormat, out reportFormat);
                var runResultsLocation = Path.Combine(TestConfiguration.TestRunSessionPath, testName);
                if (File.Exists(TestConfiguration.TestDataDrivenFile))
                {
                    testName = Utilities.GetTestListPlans()[0];
                    qtTestNameFullPath = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.UiScripts, testName);
                    TestConfiguration.RowCounter += 1;
                }
                UftTest = Open(testName);
                Thread.Sleep(5000);
                TestSettings(Settings.Run);
                TestSettings(Settings.LocalSystemMonitor);
                RunOptions(RuntimeOption.RunMode);
                if (reportFormat == ReportFormats.RRV)
                {
                    //RunOptions(RunTimeOptions.AutoExport);
                }
                if (TestConfiguration.CaptureMovie)
                {
                    RunOptions(RuntimeOption.ScreenCapture);
                }
                AddTestSourcesFolders();
                AddResources();
                Thread.Sleep(10000);
                RunResultsOptions runResults = new RunResultsOptions { ResultsLocation = runResultsLocation };
                UftTest.Run(runResults);
                Thread.Sleep(5000);
                lastRunResultStatus = Utilities.GetLastTestRunResult(testName);
                if (File.Exists(TestConfiguration.TestDataDrivenFile))
                {
                    InjectEnvironmentInfo(Path.Combine(qtTestNameFullPath, "Action1", "Script.mts"), true);
                }
            }
            catch (Exception exception)
            {
                Log.Error(exception);
            }
            return lastRunResultStatus;
        }

        public static void Execute()
        {
            IList<string> testPlans;
            var stopWatch = new Stopwatch();
            ResultStatus lastRunStatus = ResultStatus.Inconclusive;
            Log.Block("Marathon Engine", "Opened");
            Log.Block("Verify UFT Heart Beat", "Opened");
            Launch();
            Log.Block("Verify UFT Heart Beat", "Closed");
            Log.Block("Set Project WorkSpace", "Opened");
            Utilities.SetProjectWorkSpace();
            Log.Block("Set Project WorkSpace", "Closed");
            Log.Block("Get Test Plans", "Opened");
            if (File.Exists(TestConfiguration.TestDataDrivenFile))
            {
                testPlans = Utilities.GetTestListPlans(true);
            }
            else
            {
                testPlans = Utilities.GetTestListPlans();
            }
            Log.Block("Get Test Plans", "Closed");
            Log.Block("Create Test Run Session Folder", "Opened");
            Utilities.CreateTestRunFolders();
            Utilities.CreateTestRunSessionFolder();
            Log.Block("Create Test Run Session Folder", "Closed");
            if (testPlans.Count > 0)
            {
                Log.Block("Test Run Session", "Opened");
                Log.Normal("Duration      Status                Test Name");
                Log.Normal("---------------------------------------------");
                foreach (var testPlan in testPlans)
                {
                    var activeTestPlan = testPlan;
                    Parallel.Invoke
                        (
                            () =>
                            {
                                stopWatch.Start();
                                TestConfiguration.ActiveReportStack.Add(testPlan);
                                lastRunStatus = Run(activeTestPlan);
                                stopWatch.Stop();
                            },
                            () =>
                            {
                                var currentSessionReportPath = Path.Combine(TestConfiguration.TestRunSessionPath, activeTestPlan, TestConfiguration.ReportFolderName);
                                TestMonitor.WatchExecution(currentSessionReportPath);
                            }
                        );

                    var timeStamp = Utilities.GetFormattedTimeStamp(stopWatch.ElapsedMilliseconds);
                    Utilities.PrintTestRunReport(activeTestPlan, lastRunStatus.ToString(), timeStamp);
                    stopWatch.Reset();
                    if ((TestConfiguration.ConsecutiveFailureCheck.Contains("Active")) && (TestConfiguration.ConsecutiveFailureCounter == TestConfiguration.ConsecutiveMaximumFailure))
                    {
                        break;
                    }
                    ResetAutomationTool();
                }
                Log.Block("Test Run Session", "Closed");
                Log.Block("Test Report Session", "Opened");
                Utilities.PublishCrashReport();
                Utilities.GenerateTestReports(testPlans);
                Utilities.SendReportMail();
                Log.Block("Test Report Session", "Closed");
            }
            else
            {
                Log.Error("No valid test(s) found in the list. Please check your properties file.");
            }
            Log.Block("Marathon Engine", "Closed");
        }

        /// <summary>
        ///     Release all the specified process from the memory
        /// </summary>
        private static void Release()
        {
            try
            {
                var processList = TestConfiguration.KillProcessList.Split(TestConfiguration.LibrarySplitChar).ToList();
                foreach (var processName in processList)
                {
                    var isProcessExists = Process.GetProcesses().Any(p => p.ProcessName.Contains(processName));
                    if (isProcessExists)
                    {
                        //Log.Normal("Process exists. Terminating " + processName);
                        Process.GetProcesses()
                            .Where(p => p.ProcessName.Contains(processName))
                            .ToList()
                            .ForEach(p => p.Kill());
                    }
                }
            }
            catch (Exception uftException)
            {
                Log.Error(uftException);
            }
        }

        /// <summary>
        ///     Check the UFT is runnig or not
        /// </summary>
        /// <returns></returns>
        public static bool IsUftProcessExists()
        {
            var uftProcess = Process.GetProcesses().Any(p => p.ProcessName.Contains("UFT"));
            return uftProcess;
        }

        /// <summary>
        ///     Check the UFT is runnig or not
        /// </summary>
        /// <returns></returns>
        public static bool IsUftVisible()
        {
            var isUftWindowExists =
                Process.GetProcesses().Any(p => p.MainWindowTitle.Contains(TestConfiguration.UftWindowTitle));
            return isUftWindowExists;
        }

        /// <summary>
        ///     Check the UFT not responding state
        /// </summary>
        /// <returns></returns>
        public static bool IsUftNotRespondingState()
        {
            var isUftNotResponding =
                Process.GetProcesses().Any(p => p.MainWindowTitle.Contains(TestConfiguration.UftTestNotRespondingTitle));
            return isUftNotResponding;
        }

        /// <summary>
        ///     Cleanup the Uft instance
        /// </summary>
        public static void CleanUp()
        {
            UftRunOptions.Run.ReportFormat = TestConfiguration.DefaultReportFormat;
            try
            {
                if (UftTest.IsRunning == true)
                {
                    UftTest.Close();
                }
            }
            catch (Exception)
            {
                //Do nothing
            }
            finally
            {
                UftApplication.Quit();
            }
        }

        /// <summary>
        ///     Open the target uft script fil
        /// </summary>
        public static void InjectEnvironmentInfo(string scriptPath, bool remove)
        {
            string injectorLocation = Path.Combine(Path.GetTempPath(), "EnvironmentValues.xml");
            string injectorCode = "Environment.LoadFromFile(" + "\"" + injectorLocation + "\")";
            string scriptContents = File.ReadAllText(scriptPath);
            scriptContents = scriptContents.Replace(injectorCode, "");
            if (!string.IsNullOrEmpty(scriptContents) && !remove)
            {
                scriptContents = injectorCode + "\n" + scriptContents;
                File.WriteAllText(scriptPath, scriptContents);
            }
        }

        /// <summary>
        ///     Prepare the uft test scipt with data driven enabled.
        ///  </summary>
        /// <param name="variables"></param>
        public static void PrepareEnvironmentInfo(List<string> variables)
        {
            string injectorLocation = Path.Combine(Path.GetTempPath(), "EnvironmentValues.xml");
            using (XmlWriter xmlWriter = XmlWriter.Create(injectorLocation))
            {
                xmlWriter.WriteStartElement("Environment");
                foreach (string variable in variables)
                {
                    if (!string.IsNullOrEmpty(variable))
                    {
                        xmlWriter.WriteStartElement("Variable");
                        xmlWriter.WriteElementString("Name", variable.Split('=')[0]);
                        xmlWriter.WriteElementString("Value", variable.Split('=')[1]);
                        xmlWriter.WriteEndElement();
                    }
                }
                xmlWriter.WriteEndElement();
                xmlWriter.Flush();
            }
        }

        public static void ResetAutomationTool()
        {
            CrashHandler.HandleCrashWindows();
            Utilities.TerminateDeadLockBrowsers();
            Close(close: true, quitTool: true);
            UftApplication = null;
            Thread.Sleep(10000);
            Launch();
        }
    }
}