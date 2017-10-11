using System;

namespace Marathon
{
    internal class FrameworkEngine
    {
        private static readonly Logger Log = Logger.GetLogger(typeof(FrameworkEngine));

        public static void EnableExecutorMode(string[] buildArguments)
        {
            TestConfiguration.SuiteBeginTime = DateTime.Now;
            Log.Block("Read CommandLine Arguments", "Opened");
            if (buildArguments.Length == 0)
            {
                Log.Failure("Arguments list is empty. Cannot run further. Quitting now. Length is " + buildArguments.Length);
                Environment.Exit(0);
            }
            if (buildArguments.Length != 9)
            {
                Log.Error("Insufficient build arguments. Length is " + buildArguments.Length);
                Log.Normal("=================Build Arguments Supplied ===============");
                for (int i = 0; i < buildArguments.Length; i++)
                {
                    Log.Normal("Build arguments {" + i + "} =" + buildArguments[i]);
                }
                Environment.Exit(0);
            }
            TestConfiguration.ProjectWorkspace = buildArguments[0];
            TestConfiguration.FrameworkWorkspace = buildArguments[1];
            TestConfiguration.TestPropertiesFile = buildArguments[2];
            TestConfiguration.Browser = buildArguments[3];
            TestConfiguration.Environment = buildArguments[4];
            TestConfiguration.ConsecutiveFailureCheck = buildArguments[5];
            TestConfiguration.TestResourcesList = buildArguments[6];
            TestConfiguration.EmailList = buildArguments[7];
            TestConfiguration.ProjectInformation = buildArguments[8];

            Log.Normal("Configuration.ProjectWorkSpace {0} = " + buildArguments[0]);
            Log.Normal("Configuration.FrameworkWorkSpace {1} = " + buildArguments[1]);
            Log.Normal("Configuration.TestPropertiesFile {2} = " + buildArguments[2]);
            Log.Normal("Configuration.Browser {3} = " + buildArguments[3]);
            Log.Normal("Configuration.Environment {4} = " + buildArguments[4]);
            Log.Normal("Configuration.ConsecutiveFailureCheck {5} = " + buildArguments[5]);
            Log.Normal("Configuration.TestResourcesList {6} = " + buildArguments[6]);
            Log.Normal("Configuration.EmailList {7} = " + buildArguments[7]);
            Log.Normal("Configuration.ProjectInformation {8} = " + buildArguments[8]);
        }
    }
}