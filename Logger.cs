using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace Marathon
{
    public class Logger
    {
        private string mTypeName;

        static Logger()
        {
            LogLevel = LogLevel.Normal;
        }

        private Logger(string typeName)
        {
            mTypeName = typeName;
        }

        public static LogLevel LogLevel { get; set; }

        public static Logger GetLogger(Type type)
        {
            return new Logger(type.Name);
        }

        public void Normal(string logText, params object[] args)
        {
            if (LogLevel <= LogLevel.Normal)
            {
                AddToLog(logText, LogLevel.Normal);
            }
        }

        public void Warning(string logText, params object[] args)
        {
            if (LogLevel <= LogLevel.Warning)
            {
                AddToLog(logText, LogLevel.Warning);
            }
        }

        public void Failure(string logText, params object[] args)
        {
            if (LogLevel <= LogLevel.Failure)
            {
                AddToLog(logText, LogLevel.Failure);
            }
        }

        public void Error(string logText, params object[] args)
        {
            if (LogLevel <= LogLevel.Error)
            {
                AddToLog(logText, LogLevel.Error);
            }
        }

        public void Block(string blockName, string blockState)
        {
            if (LogLevel <= LogLevel.Error)
            {
                AddToBlockLog(blockName, blockState);
            }
        }

        public void Error(Exception ex)
        {
            if (LogLevel <= LogLevel.Error)
            {
                AddToLog(ex.Message, LogLevel.Failure);
            }
        }

        private void AddToBlockLog(string logText, string logState)
        {
            Console.WriteLine(logText);
        }

        private void AddToLog(string format, LogLevel logLevel)
        {
            Console.WriteLine(format);
        }

        /// <summary>
        ///     Saves a picture of the screen to a bitmap image.
        /// </summary>
        /// <returns>The saved bitmap.</returns>
        public static string CaptureScreenshot()
        {
            string currentTestName = UnifiedFunctionalTesting.UftTest.Name;
            string timeStamp = "Crash_" + currentTestName +
                               DateTime.Now.ToString("HHmmssffff") + "." + ImageFormat.Png;
            string fileName = Path.Combine(TestConfiguration.CrashReportPath, timeStamp);
            var bounds = Screen.GetBounds(Point.Empty);
            using (var bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (var gr = Graphics.FromImage(bitmap))
                {
                    gr.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }
                bitmap.Save(fileName);
            }
            return timeStamp;
        }
    }
}