using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using System.Management;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Marathon
{
    internal class Utilities
    {
        private static readonly Logger Log = Logger.GetLogger(typeof(Utilities));

        /// <summary>
        ///     Transfom result xml file into html file using the custom stylesheet.
        /// </summary>
        /// <param name="xmlFilePath"> report xml file path</param>
        /// <param name="xslFilePath">custom stylesheet file path</param>
        /// <param name="htmlFilePath">html report file path</param>
        private static void CreateHtmlReport(string xmlFilePath, string xslFilePath, string htmlFilePath)
        {
            try
            {
                XmlWriter xmlWriter = XmlWriter.Create(htmlFilePath);
                XmlReaderSettings xmlReaderSettings = new XmlReaderSettings();
                xmlReaderSettings.XmlResolver = null;
                xmlReaderSettings.DtdProcessing = DtdProcessing.Ignore;
                var xslTransform = new XslCompiledTransform();
                xslTransform.Load(xslFilePath);
                xslTransform.Transform(XmlReader.Create(xmlFilePath, xmlReaderSettings), xmlWriter);
                xmlWriter.Close();
            }
            catch (Exception exception)
            {
                Log.Error(exception);
            }
        }

        /// <summary>
        /// Create a run session folders
        /// </summary>
        public static void CreateTestRunFolders()
        {
            var htmlReportPath = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.HtmlReport);
            var resultsReportPath = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.Results);
            var failureSnapsReportPath = Path.Combine(htmlReportPath, TestConfiguration.Screenshots);
            TestConfiguration.CrashReportPath = failureSnapsReportPath;
            List<string> folderList = new List<string> { htmlReportPath, resultsReportPath, failureSnapsReportPath };
            try
            {
                foreach (string folder in folderList)
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(folder);
                    if (directoryInfo.Exists)
                    {
                        directoryInfo.Delete(true);
                    }
                    Directory.CreateDirectory(folder);
                }
                TestConfiguration.HtmlScreenshotsFolder = failureSnapsReportPath;
                Log.Normal("HTML folder created as => " + htmlReportPath);
                Log.Normal("Results folder created as => " + resultsReportPath);
                Log.Normal("Failure snapshot folder created as => " + failureSnapsReportPath);
            }
            catch (DirectoryNotFoundException directoryNotFoundException)
            {
                Log.Error("Unable to deleting the directory => " + htmlReportPath);
                Log.Error(directoryNotFoundException);
            }
        }

        /// <summary>
        ///     Create a test run session folders
        /// </summary>
        public static void CreateTestRunSessionFolder()
        {
            var resultsReportPath = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.Results);
            var timeStamp = string.Format(CultureInfo.CurrentCulture, "{0:yyyy-MM-dd_hh-mm-ss-tt}", DateTime.Now);
            var testRunsPath = Path.Combine(resultsReportPath, timeStamp);
            if (!Directory.Exists(testRunsPath))
            {
                Directory.CreateDirectory(testRunsPath);
                Log.Normal("Test Run Session Folder Created", testRunsPath);
                TestConfiguration.TestRunSessionPath = testRunsPath;
            }
        }

        /// <summary>
        ///     Get the formatted time stamp
        /// </summary>
        /// <param name="totalMilliSeconds">Time elapsed </param>
        /// <returns>HH:MM:SS</returns>
        public static string GetFormattedTimeStamp(double totalMilliSeconds)
        {
            var timeSpan = TimeSpan.FromMilliseconds(totalMilliSeconds);
            return string.Format(CultureInfo.CurrentCulture, "{0:D2}:{1:D2}:{2:D2}", timeSpan.Hours,
                timeSpan.Minutes, timeSpan.Seconds);
        }

        /// <summary>
        ///     Get the last runs status of the completed test suite.
        /// </summary>
        /// <param name="testPlan"></param>
        /// <returns></returns>
        public static ResultStatus GetLastTestRunResult(string testPlan)
        {
            var result = ResultStatus.Inconclusive;
            string xmlReportPath;
            ReportFormats reportFormats;
            Enum.TryParse(TestConfiguration.ReportFormat, out reportFormats);
            if (reportFormats == ReportFormats.HTML)
            {
                xmlReportPath = Path.Combine(TestConfiguration.TestRunSessionPath, testPlan, TestConfiguration.ReportFolderName, TestConfiguration.HtmlReportXmlFileName);
            }
            else
            {
                xmlReportPath = Path.Combine(TestConfiguration.TestRunSessionPath, testPlan, TestConfiguration.ReportFolderName, TestConfiguration.RrvReportXmlFileName);
            }
            if (File.Exists(xmlReportPath))
            {
                var xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlReportPath);
                string TestSuiteResultNodePath = "null";
                if (TestConfiguration.ReportFormat == "RRV")
                {
                    TestSuiteResultNodePath = TestConfiguration.RrvTestSuiteResultNodePath;
                }
                else
                {
                    TestSuiteResultNodePath = TestConfiguration.HtmlTestSuiteResultNodePath;
                }

                var xmlNodeList = xmlDocument.SelectNodes(TestSuiteResultNodePath);
                if (xmlNodeList != null)
                    foreach (XmlNode xmlNode in xmlNodeList)
                    {
                        var resultStatus = xmlNode.InnerText;
                        Enum.TryParse(resultStatus, out result);
                        break;
                    }
            }
            if (result == ResultStatus.Warning)
            {
                result = ResultStatus.Passed;
            }
            return result;
        }

        /// <summary>
        ///     Print the build arguments list
        /// </summary>
        /// <param name="buildArguments"></param>
        public static void PrintArgumentList(string[] buildArguments)
        {
            for (var i = 0; i < buildArguments.Length; i++)
            {
                Log.Normal(i + "=" + buildArguments[i]);
            }
        }

        /// <summary>
        ///     Print the test report.
        /// </summary>
        /// <param name="testPlan">Name of test plan scheduled</param>
        /// <param name="runStatus">Status of the run</param>
        /// <param name="timeStamp">Time stamp</param>
        public static void PrintTestRunReport(string testPlan, string runStatus, string timeStamp)
        {
            switch (runStatus.ToUpper())
            {
                case "PASSED":
                    Log.Normal(timeStamp + "      " + runStatus + "        " + testPlan);
                    break;

                case "FAILED":
                case "ERROR":
                    Log.Failure(timeStamp + "      " + runStatus + "        " + testPlan);
                    TestConfiguration.ConsecutiveMaximumFailure += 1;
                    break;

                case "WARNING":
                case "SKIPPED":
                    Log.Warning(timeStamp + "      " + runStatus + "        " + testPlan);
                    break;

                default:
                    Log.Normal(timeStamp + "      " + runStatus + "        " + testPlan);
                    break;
            }
            TestConfiguration.UftTestResults.Add(testPlan + "=" + runStatus.ToUpper());
        }

        /// <summary>
        ///     Publish the test reports
        /// </summary>
        /// <param name="testPlans"></param>
        public static void GenerateTestReports(IList<string> testPlans)
        {
            try
            {
                TestConfiguration.RowSequence = 1;
                ReportFormats reportFormats;
                Enum.TryParse(TestConfiguration.ReportFormat, out reportFormats);
                var frameworkHtmlFolder = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.HtmlReport);
                var summaryReportName = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.HtmlReport, TestConfiguration.RunStatistics);
                CreateTestSummaryHeader(summaryReportName);
                foreach (var testPlan in testPlans)
                {
                    var resultFolderPath = Path.Combine(TestConfiguration.TestRunSessionPath, testPlan, TestConfiguration.ReportFolderName);
                    if (reportFormats == ReportFormats.HTML)
                    {
                        string defaultResultsPath = Path.Combine(resultFolderPath, TestConfiguration.HtmlReportXmlFileName);
                        string detailedHtmlReport = Path.Combine(resultFolderPath, TestConfiguration.HtmlReportHtmlFileName);
                        var xsltReportPath = Path.Combine(TestConfiguration.FrameworkWorkspace, TestConfiguration.ReportXslHtml);
                        var customizedHtmlReport = Path.Combine(frameworkHtmlFolder, testPlan + ".htm");
                        var detailedHtmlReportTestPlan = Path.Combine(frameworkHtmlFolder, testPlan + ".html");
                        CreateHtmlReport(defaultResultsPath, xsltReportPath, customizedHtmlReport);
                        CopyTestReport(detailedHtmlReport, detailedHtmlReportTestPlan);
                        var defaultRunMovie = Path.Combine(resultFolderPath, TestConfiguration.RrvReportMovieName);
                        var defaultRunMovieTestPlan = Path.Combine(frameworkHtmlFolder, testPlan + ".fbr");
                        CopyTestReport(defaultRunMovie, defaultRunMovieTestPlan);
                        GenerateTestStatistics(testPlan, defaultResultsPath);
                    }
                    else
                    {
                        string defaultResultsPath = Path.Combine(resultFolderPath, TestConfiguration.RrvReportXmlFileName);
                        string customizedHtmlReport = Path.Combine(frameworkHtmlFolder, testPlan + ".htm");
                        string detailedHtmlReport = Path.Combine(frameworkHtmlFolder, testPlan + ".html");
                        string xsltReportPath = Path.Combine(TestConfiguration.FrameworkWorkspace, TestConfiguration.ReportXslRrv);
                        CreateHtmlReport(defaultResultsPath, xsltReportPath, customizedHtmlReport);
                        xsltReportPath = Path.Combine(TestConfiguration.FrameworkWorkspace, TestConfiguration.ReportXslRrvDetailed);
                        CreateHtmlReport(defaultResultsPath, xsltReportPath, detailedHtmlReport);
                        var defaultRunMovie = Path.Combine(resultFolderPath, TestConfiguration.RrvReportMovieName);
                        var defaultRunMovieTestPlan = Path.Combine(frameworkHtmlFolder, testPlan + ".fbr");
                        CopyTestReport(defaultRunMovie, defaultRunMovieTestPlan);
                        GenerateTestStatistics(testPlan, defaultResultsPath);
                    }
                    TestConfiguration.RowSequence += 1;
                }
            }
            catch (Exception exception)
            {
                Log.Normal(exception.Message);
            }
        }

        /// <summary>
        ///     Get the test plans scheduled for the current run.
        /// </summary>
        /// <returns></returns>
        protected internal static List<string> GetTestListPlans()
        {
            var qtActiveTests = new List<string>();
            if (File.Exists(TestConfiguration.TestPropertiesFile))
            {
                var testPlans = File.ReadAllLines(TestConfiguration.TestPropertiesFile);
                if (!File.Exists(TestConfiguration.TestDataDrivenFile))
                {
                    Log.Normal("List of test plans available");
                    Log.Normal("----------------------------");
                }
                foreach (var testPlan in testPlans)
                {
                    char commentChar = testPlan[0];
                    if (commentChar != TestConfiguration.CommentCharacter)
                    {
                        Log.Normal(testPlan);
                        qtActiveTests.Add(testPlan.Trim());
                    }
                }
            }
            return qtActiveTests;
        }

        /// <summary>
        ///     Get the test plans scheduled for the current run for data driven approach.
        /// </summary>
        /// <returns></returns>
        protected internal static List<string> GetTestListPlans(bool datadriven)
        {
            var qtActiveTests = new List<string>();
            if (File.Exists(TestConfiguration.TestDataDrivenFile))
            {
                Application excelApp = new Application();
                Workbook excelWorkBook = excelApp.Workbooks.Open(TestConfiguration.TestDataDrivenFile);
                Worksheet excelWorkSheet = GetDesiredWorkSheet(excelApp);
                Range excelRange = excelWorkSheet.UsedRange;
                int rows = excelRange.Rows.Count;
                int j = 1;
                if (rows > 0)
                {
                    Log.Normal("List of test plans available");
                    Log.Normal("----------------------------");
                    for (int i = 1; i <= rows; i++)
                    {
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        {
                            bool include = excelRange.Cells[i, j].Value2.ToString().Trim().ToUpper() == "YES";
                            if (include)
                            {
                                string testPlan = excelRange.Cells[i, j + 1].Value2;
                                Log.Normal(testPlan);
                                qtActiveTests.Add(testPlan.Trim());
                            }
                        }
                    }
                }
                Marshal.ReleaseComObject(excelRange);
                Marshal.ReleaseComObject(excelWorkSheet);
                excelWorkBook.Close();
                Marshal.ReleaseComObject(excelWorkBook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            return qtActiveTests;
        }

        private static Worksheet GetDesiredWorkSheet(Application excelApp)
        {
            Worksheet targetSheet = null;
            Sheets availableWorkSheets = excelApp.Sheets;
            foreach (Worksheet sheet in availableWorkSheets)
            {
                if (sheet.Name.Trim().ToUpper() == TestConfiguration.Environment)
                {
                    targetSheet = sheet;
                    break;
                }
            }
            return targetSheet;
        }

        /// <summary>
        ///     Set the project workspace path calculated
        /// </summary>
        protected internal static void SetProjectWorkSpace()
        {
            Log.Normal("Project WorkSpace =>" + TestConfiguration.ProjectWorkspace);
            Log.Normal("Framework WorkSpace =>" + TestConfiguration.FrameworkWorkspace);
        }

        /// <summary>
        ///     Copy test reports in to new file
        /// </summary>
        /// <param name="oldFileName">Old file name</param>
        /// <param name="newFileName">new file name</param>
        private static void CopyTestReport(string oldFileName, string newFileName)
        {
            try
            {
                File.Copy(oldFileName, newFileName);
            }
            catch (AccessViolationException accessViolationException)
            {
                Log.Error(accessViolationException);
            }
            catch (FileNotFoundException fileNotFoundException)
            {
                Log.Error(fileNotFoundException);
            }
            catch (Exception exception)
            {
                Log.Error(exception);
            }
        }

        /// <summary>
        ///     Create the test summary report
        /// </summary>
        /// <param name="testName"></param>
        /// <param name="runResultsXml"></param>
        public static void GenerateTestStatistics(string testName, string runResultsXml)
        {
            ReportFormats reportFormats;
            XmlNodeList xmlNodeList;
            XmlNodeList nodes;
            double duration = 0;
            Enum.TryParse(TestConfiguration.ReportFormat, out reportFormats);
            var runStats = new Dictionary<string, string>();
            var passCount = 0;
            var failCount = 0;
            var warningCount = 0;
            var runStatus = "PASSED";
            var xmlDocument = new XmlDocument();
            xmlDocument.Load(runResultsXml);
            if (reportFormats == ReportFormats.HTML)
            {
                xmlNodeList = xmlDocument.SelectNodes("/Results/ReportNode/ReportNode/ReportNode/Data/Result");
            }
            else
            {
                xmlNodeList = xmlDocument.SelectNodes("/Report/Doc/Action/Step/NodeArgs/@status");
            }
            if (xmlNodeList != null)
                foreach (XmlNode xmlNode in xmlNodeList)
                {
                    var stepResult = xmlNode.InnerText.ToUpper();
                    switch (stepResult)
                    {
                        case "PASSED":
                            passCount += 1;
                            break;

                        case "FAILED":
                            failCount += 1;
                            break;

                        case "WARNING":
                            warningCount += 1;
                            break;
                    }
                }
            if (failCount > 0)
            {
                runStatus = "FAILED";
            }
            else if (warningCount > 0)
            {
                runStatus = "WARNING";
            }
            if (runStatus == "PASSED")
            {
                TestConfiguration.Passed += 1;
            }
            else
            {
                TestConfiguration.Failed += 1;
            }
            if (reportFormats == ReportFormats.HTML)
            {
                nodes = xmlDocument.SelectNodes(TestConfiguration.HtmlReportDurationPath);
                if (nodes != null) duration = int.Parse(nodes[0].InnerText) * 1000;
            }
            else
            {
                string startDateTimeStamp = xmlDocument.SelectNodes(TestConfiguration.RrvlReportDurationStartTimePath)?[0].InnerText;
                string endDateTimeStamp = xmlDocument.SelectNodes(TestConfiguration.RrvlReportDurationEndTimePath)?[0].InnerText;
                startDateTimeStamp = startDateTimeStamp?.Replace("- ", "");
                endDateTimeStamp = endDateTimeStamp?.Replace("- ", "");
                string startDateStamp = FormatDateTimeStamp(startDateTimeStamp.Split(' ')[0], '/');
                string startTimeStamp = FormatDateTimeStamp(startDateTimeStamp.Split(' ')[1], ':');
                string endDateStamp = FormatDateTimeStamp(endDateTimeStamp.Split(' ')[0], '/');
                string endTimeStamp = FormatDateTimeStamp(endDateTimeStamp.Split(' ')[1], ':');
                try
                {
                    DateTime suiteStartTime = DateTime.ParseExact(startDateStamp + " " + startTimeStamp, "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    DateTime suiteEndTime = DateTime.ParseExact(endDateStamp + " " + endTimeStamp, "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    TimeSpan timeElapsed = suiteEndTime - suiteStartTime;
                    duration = timeElapsed.TotalMilliseconds;
                }
                catch (Exception exception)
                {
                    Log.Warning("Unable to parse the start and end time. Exception raised." + exception.Message);
                }
            }
            runStats.Add("Sequence", TestConfiguration.RowSequence.ToString());
            runStats.Add("Name", testName);
            runStats.Add("Status", runStatus);
            runStats.Add("StepPassed", passCount.ToString());
            runStats.Add("StepFailed", failCount.ToString());
            if (failCount == 0)
            {
                TestConfiguration.PassedTests.Add(testName);
            }
            runStats.Add("StepWarning", warningCount.ToString());
            runStats.Add("Duration", GetFormattedTimeStamp(duration).Replace("-", ""));
            if (TestConfiguration.TriggerMode == Trigger.Local)
            {
                string frameworkHtmlFolder = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.HtmlReport);
                runStats.Add("StepReport", Path.Combine(frameworkHtmlFolder, testName + ".htm"));
                runStats.Add("Report", Path.Combine(frameworkHtmlFolder, testName + ".html"));
                if (TestConfiguration.CaptureMovie)
                {
                    runStats.Add("Movie", Path.Combine(frameworkHtmlFolder, testName + ".fbr"));
                }
                else
                {
                    runStats.Add("Movie", "");
                }
            }
            CreateTestSummary(runStats);
            runStats.Clear();
        }

        public static void PublishCrashReport()
        {
            Log.Block("Crash Reports", "Opened");
            Log.Normal("Crash Report Files");
            Log.Normal("-------------------");
            foreach (string crashFile in TestConfiguration.CrashListTable)
            {
                string crashReportLink = "Crash SnapShots:" + crashFile;
                Log.Normal(crashReportLink);
            }
            Log.Block("Crash Reports", "Closed");
        }

        public static string FormatDateTimeStamp(string givenDateTimeStamp, char splitChar)
        {
            string formattedTimeStamp = string.Empty;
            try
            {
                var timeStamps = givenDateTimeStamp.Split(splitChar).ToList();
                foreach (string timeStamp in timeStamps)
                {
                    if (timeStamp.Length == 1)
                    {
                        formattedTimeStamp = formattedTimeStamp + splitChar + "0" + timeStamp;
                    }
                    else
                    {
                        formattedTimeStamp = formattedTimeStamp + splitChar + timeStamp;
                    }
                }
            }
            catch (Exception)
            {
                //Not important at this moment
            }
            return formattedTimeStamp.Substring(1);
        }

        public static void CreateTestSummaryHeader(string summaryReportFile)
        {
            BrowserType browserTypes;
            TestConfiguration.SuiteEndTime = DateTime.Now;
            TimeSpan suiteTimeElapsed = TestConfiguration.SuiteEndTime - TestConfiguration.SuiteBeginTime;
            string suiteTimeTaken = GetFormattedTimeStamp(suiteTimeElapsed.TotalMilliseconds);
            Enum.TryParse(TestConfiguration.BrowserType, out browserTypes);
            string hostname = Environment.MachineName;
            TimeZoneInfo localZone = TimeZoneInfo.Local;
            string displayName = localZone.DisplayName;
            if (displayName != string.Empty)
            {
                displayName = displayName.Split(' ')[0].Replace("(", "").Replace(")", "");
            }
            string timezone = displayName;
            string userName = GetLoggedUserName();
            if (userName == string.Empty)
            {
                userName = Environment.GetEnvironmentVariable("UserName");
            }
            string username = Environment.UserDomainName + "\\" + userName;
            string browser = GetBrowserVersion(browserTypes);
            string productname = TestConfiguration.UftWindowTitle + " " + UnifiedFunctionalTesting.UftApplication.Version;
            if (!File.Exists(summaryReportFile))
            {
                File.Delete(summaryReportFile);
            }
            var xmlWriterSettings = new XmlWriterSettings
            {
                Indent = true,
                NewLineOnAttributes = true
            };
            using (var xmlWriter = XmlWriter.Create(summaryReportFile, xmlWriterSettings))
            {
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("Environment");
                xmlWriter.WriteStartElement("Summary");
                xmlWriter.WriteElementString("HostName", hostname);
                xmlWriter.WriteElementString("Starttime", TestConfiguration.SuiteBeginTime.ToString());
                xmlWriter.WriteElementString("Timezone", timezone);
                xmlWriter.WriteElementString("Endtime", TestConfiguration.SuiteEndTime.ToString());
                xmlWriter.WriteElementString("Browser", browser);
                xmlWriter.WriteElementString("Elapsed", suiteTimeTaken);
                xmlWriter.WriteElementString("User", username);
                xmlWriter.WriteElementString("Version", productname);
                xmlWriter.WriteElementString("Project", TestConfiguration.ProjectInformation.Split('|')[0]);
                xmlWriter.WriteElementString("Component", TestConfiguration.ProjectInformation.Split('|')[1]);
                xmlWriter.WriteElementString("SuiteType", TestConfiguration.ProjectInformation.Split('|')[2]);
                xmlWriter.WriteElementString("RunEnvironment", TestConfiguration.ProjectInformation.Split('|')[3]);
                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndDocument();
                xmlWriter.Flush();
                xmlWriter.Close();
            }
        }

        public static string GetBrowserVersion(BrowserType browserType)
        {
            string Version = string.Empty;
            switch (browserType)
            {
                case BrowserType.IE:
                    using (
                        RegistryKey key =
                            Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Internet Explorer"))
                    {
                        if (key != null)
                        {
                            Version = "Internet Explorer - " + key.GetValue("svcVersion").ToString();
                        }
                    }
                    break;

                case BrowserType.Firefox:
                    break;

                case BrowserType.Chrome:
                    break;
            }
            return Version;
        }

        /// <summary>
        ///     Create the test summary
        /// </summary>
        /// <param name="runStatistics"></param>
        public static void CreateTestSummary(Dictionary<string, string> runStatistics)
        {
            IEnumerable<XElement> childNode = null;
            var summaryReportName = Path.Combine(TestConfiguration.ProjectWorkspace, TestConfiguration.HtmlReport,
                TestConfiguration.RunStatistics);
            var xmlDocument = XDocument.Load(summaryReportName);
            XElement root = xmlDocument.Element("Environment");
            if (root != null)
            {
                childNode = root.Descendants("Test");
                if (childNode.Count() == 0)
                {
                    childNode = root.Descendants("Summary");
                }
            }
            XElement lastRow = childNode.Last();
            if (lastRow != null)
            {
                if (runStatistics["StepFailed"] == "0")
                {
                    runStatistics["Movie"] = "";
                }
                lastRow.AddAfterSelf(
                    new XElement("Test",
                        new XElement("Sequence", runStatistics["Sequence"]),
                        new XElement("Name", runStatistics["Name"]),
                        new XElement("Status", runStatistics["Status"].ToUpper(CultureInfo.CurrentCulture)),
                        new XElement("StepPassed", runStatistics["StepPassed"]),
                        new XElement("StepFailed", runStatistics["StepFailed"]),
                        new XElement("StepWarning", runStatistics["StepWarning"]),
                        new XElement("Duration", runStatistics["Duration"]),
                        new XElement("StepReport", runStatistics["StepReport"]),
                        new XElement("Report", runStatistics["Report"]),
                        new XElement("Movie", runStatistics["Movie"])));
                xmlDocument.Save(summaryReportName);
            }
        }

        /// <summary>
        /// Get the logged in user name
        /// </summary>
        /// <returns></returns>
        private static string GetLoggedUserName()
        {
            string loggedinUserName = string.Empty;
            try
            {
                ManagementObjectSearcher processes = new ManagementObjectSearcher("SELECT * FROM Win32_Process");
                foreach (ManagementObject process in processes.Get())
                {
                    if (process["ExecutablePath"] != null &&
                        Path.GetFileName(process["ExecutablePath"].ToString()).ToLower() == "explorer.exe")
                    {
                        string[] ownerInfo = new string[2];
                        process.InvokeMethod("GetOwner", (object[])ownerInfo);
                        loggedinUserName = ownerInfo[0];
                        break;
                    }
                }
            }
            catch
            {
                // ignored
            }
            return loggedinUserName;
        }

        /// <summary>
        /// Send Report Mail to Stakeholders
        /// </summary>
        public static void SendReportMail()
        {
            using (new MailMessage())
            {
                string mailAddress = TestConfiguration.EmailList;
                if (mailAddress.Contains(';'))
                {
                    mailAddress = mailAddress.Replace(';', ',');
                }
                var mailMessage = new MailMessage(TestConfiguration.EmailSenderAddress, mailAddress);
                SmtpClient smtpClient;
                using (smtpClient = new SmtpClient
                {
                    UseDefaultCredentials = false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Host = TestConfiguration.EmailHostAddress,
                    Port = 25
                })
                {
                    mailMessage.Subject = TestConfiguration.ProjectInformation.Split('|')[0] + " | " +
                                          TestConfiguration.ProjectInformation.Split('|')[2] + " | " +
                                          TestConfiguration.ProjectInformation.Split('|')[3] + " | Total = " +
                                          (TestConfiguration.Passed + TestConfiguration.Failed) +
                                          " | Passed = " +
                                          TestConfiguration.Passed +
                                          " | Failed = " + TestConfiguration.Failed;
                    mailMessage.IsBodyHtml = true;
                    mailMessage.Body = File.ReadAllText(TestConfiguration.TestStatisticsReportName);
                    string mailBody = mailMessage.Body;
                    mailMessage.Body = mailBody;
                    smtpClient.Send(mailMessage);
                }
            }
        }

        public static bool GetRespondingState(string processName)
        {
            bool isAlive = false;
            Process[] listOfProcesses = Process.GetProcesses();
            foreach (Process runningProcess in listOfProcesses)
            {
                if (runningProcess.ProcessName.ToUpper().Trim() == processName.ToUpper().Trim())
                {
                    isAlive = runningProcess.Responding;
                    break;
                }
            }
            return isAlive;
        }

        /// <summary>
        /// Terminate the possible deadlock browsers.
        /// </summary>
        public static void TerminateDeadLockBrowsers()
        {
            Process[] listOfProcess = Process.GetProcesses();
            List<string> deadlockBrowsers = TestConfiguration.PossibleDeadockProcesslist.Split(',').ToList();
            foreach (Process runningProcess in listOfProcess)
            {
                foreach (string deadlockBrowser in deadlockBrowsers)
                {
                    if (runningProcess.ProcessName.Trim().ToUpper() == deadlockBrowser.Trim().ToUpper())
                    {
                        try
                        {
                            runningProcess.Kill();
                            runningProcess.WaitForExit();
                        }
                        catch (Exception e)
                        {
                            Log.Error(e);
                        }
                    }
                }
            }
        }
    }
}