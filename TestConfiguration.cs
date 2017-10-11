using System;
using System.Collections.Generic;
using System.Configuration;

namespace Marathon
{
    public static class TestConfiguration
    {
        //Iteration Mode
        public static readonly string IterationMode = "oneIteration";
        //Report Folder Name
        public static readonly string HtmlReport = "HTMLReport";
        public static readonly string UiScripts = "<Your Automation Scripts Folder>";
        public static readonly string Results = "Results";
        public static readonly string Screenshots = "ScreenShots";
        public static readonly bool ViewResults = false;
        public static readonly string ImageCaptureForTestResults = "Always";
        public static readonly string MovieCaptureForTestResults = "Always";
        public static readonly int MovieSegmentSize = 2048;
        public static readonly bool SaveMovieOfEntireRun = true;
        public static readonly int StepExecutionDelay = 0;
        public static readonly string StepDetailsReportType = "HTML";
        public static readonly string HtmlTestSuiteResultNodePath = "/Results/ReportNode/Data/Result";
        public static readonly string RrvTestSuiteResultNodePath = "/Report/Doc/NodeArgs/@status";
        public static readonly string HtmlReportDurationPath = "/Results/ReportNode/Data/Duration";
        public static readonly string RrvlReportDurationStartTimePath = "/Report/Doc/Summary/@eTime";
        public static readonly string RrvlReportDurationEndTimePath = "/Report/Doc/Summary/@sTime";
        public static readonly string RunStatistics = "RunStatistics.xml";
        public static readonly string MainReportHtml = "index.html";
        public static readonly string RunStatisticsStyleSheet = "ReportBuilder_Summary.xsl";
        public static readonly string ReportXslHtml = "ReportBuilder_HTML.xsl";
        public static readonly string ReportXslRrv = "ReportBuilder_RRV.xsl";
        public static readonly string ReportXslRrvDetailed = "ReportBuilder_RRVDetailed.xsl";
        public static readonly string RrvReportXmlFileName = "Results.xml";
        public static readonly string HtmlReportXmlFileName = "run_results.xml";
        public static readonly string HtmlReportHtmlFileName = "run_results.html";
        public static readonly string HtmlReportMovieName = "run_movie.fbr";
        public static readonly string RrvReportMovieName = "MSR.fbr";
        public static readonly string ReportFolderName = "Report";
        public static readonly char LibrarySplitChar = ',';
        public static readonly string CommonLibsFolder = "<Your Framework Folder>";
        public static readonly string TestDataFolder = "<Your Test Data Folder>";
        public static readonly string UftWindowTitle = "HPE Unified Functional Testing";
        public static readonly string UftTestNotRespondingTitle = "(NotResponding)";
        public static readonly string LibraryFolder = "<Your Application Library Folder>";
        public static readonly string DefaultSpreadsheet = "Default.xls";
        public static readonly char CommentCharacter = '\'';
        public static readonly string HtmlReportEnvironmentPath = "/Results/ReportNode/Data/Environment";
        public static readonly MSO_VERSION OFFICE_VERSION;
        public static readonly string ExcelRegistryPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe";
        public static readonly string EmailSenderAddress = "automation.runs@yourdomain.com";
        public static readonly string EmailHostAddress = "Host Email Provider";
        static TestConfiguration()
        {
            Initialize();
        }
        public static int RowSequence { get; set; }
        public static string ReportFormat { get; set; }
        public static string DefaultReportFormat { get; set; }
        public static bool ShowAutomationTool { get; set; }
        public static string ProjectWorkspace { get; set; }
        public static string KillProcessList { get; set; }
        public static string TestRunSessionPath { get; set; }
        public static string OnError { get; set; }
        public static int ObjectSyncTimeout { get; set; }
        public static string MonitorApplicationProcess { get; set; }
        public static string RunMode { get; set; }
        public static string TestPropertiesFile { get; set; }
        public static string TestResourcesList { get; set; }
        public static string CommonLibsTestResourcesList { get; set; }
        public static string ProjectInformation { get; set; }
        public static string ConsecutiveFailureCheck { get; set; }
        public static string FrameworkWorkspace { get; set; }
        public static string CrashProcess { get; set; }
        public static DateTime SuiteBeginTime { get; set; }
        public static DateTime SuiteEndTime { get; set; }
        public static int TestRunSessionTimeout { get; set; }
        public static string RunMouseByEvents { get; set; }
        public static string MailList { get; set; }
        public static string ApplicationName { get; set; }
        public static string SuiteType { get; set; }
        public static int Passed { get; set; }
        public static int Failed { get; set; }
        public static string TestStatisticsReportName { get; set; }
        public static string HtmlScreenshotsFolder { get; set; }
        public static IList<string> CrashListTable { get; internal set; }
        public static IList<string> UftTestResults { get; internal set; }
        public static IList<string> PassedTests { get; internal set; }
        public static string CrashReportPath { get; set; }
        public static string BaseDirectory { get; set; }
        public static string BrowserType { get; set; }
        public static bool CaptureMovie { get; set; }
        public static string EmailList { get; set; }
        public static int ConsecutiveFailureCounter { get; set; }
        public static string Environment { get; set; }
        public static string Locale { get; set; }
        public static string TargetName { get; set; }
        public static string Browser { get; set; }
        public static string AddIns { get; set; }
        public static Trigger TriggerMode { get; set; }
        public static int ConsecutiveMaximumFailure { get; set; }
        public static string TestDataDrivenSheet { get; set; }
        public static string TestDataDrivenFile { get; set; }
        public static string AdditionalAddIns { get; set; }
        public static int RowCounter { get; set; }
        public static string PossibleDeadockProcesslist { get; set; }
        public static List<string> ActiveReportStack { get; internal set; }
        public static bool OverWritedefault { get; set; }

        private static void Initialize()
        {
            ReportFormat = ConfigurationManager.AppSettings["ReportFormat"];
            DefaultReportFormat = ConfigurationManager.AppSettings["DefaultReportFormat"];
            ShowAutomationTool = Convert.ToBoolean(ConfigurationManager.AppSettings["ShowAutomationTool"]);
            OnError = ConfigurationManager.AppSettings["OnError"];
            ObjectSyncTimeout = int.Parse(ConfigurationManager.AppSettings["ObjectSyncTimeOut"]);
            MonitorApplicationProcess = ConfigurationManager.AppSettings["MonitorApplicationProcess"];
            RunMode = ConfigurationManager.AppSettings["RunMode"];
            KillProcessList = ConfigurationManager.AppSettings["KillProcessList"];
            CommonLibsTestResourcesList = ConfigurationManager.AppSettings["CommonLibsTestResourcesList"];
            TestRunSessionTimeout = int.Parse(ConfigurationManager.AppSettings["TestRunSessionTimeOut"]);
            RunMouseByEvents = ConfigurationManager.AppSettings["RunMouseByEvents"];
            CrashProcess = ConfigurationManager.AppSettings["CrashProcess"];
            CaptureMovie = Convert.ToBoolean(ConfigurationManager.AppSettings["CaptureMovie"]);
            AddIns = ConfigurationManager.AppSettings["AddIns"];
            EmailList = ConfigurationManager.AppSettings["EmailList"];
            ConsecutiveFailureCounter = int.Parse(ConfigurationManager.AppSettings["ConsecutiveFailureCounter"]);
            PossibleDeadockProcesslist = ConfigurationManager.AppSettings["PossibleDeadLockProcessList"];
            UftTestResults = new List<string>();
            CrashListTable = new List<string>();
            PassedTests = new List<string>();
            ActiveReportStack = new List<string>();
            RowCounter = 1;
            RowSequence = 0;
            OverWritedefault = false;
        }
    }

    public enum LogLevel
    {
        Normal = 0,
        Warning,
        Failure,
        Error
    }

    public enum ResultStatus
    {
        Passed = 0,
        Failed,
        Warning,
        Error,
        Inconclusive
    }

    public enum Settings
    {
        Run = 0,
        Recovery,
        LocalSystemMonitor,
        Environment
    }

    public enum RuntimeOption
    {
        RunMode = 0,
        ScreenCapture,
        WebAdvanced,
        AutoExport,
        Folders
    }

    public enum ReportFormats
    {
        RRV = 0,
        HTML
    }

    public enum BrowserType
    {
        IE = 0,
        Firefox,
        Chrome
    }

    public enum Trigger
    {
        Local = 0,
        Remote
    }

    public enum MSO_VERSION
    {
        MSO_NOTSUPPORTED = 0,
        MSO2007 = 2007,
        MSO2010 = 2010,
        MSO2013 = 2013,
        MSO2016 = 2016
    }
}