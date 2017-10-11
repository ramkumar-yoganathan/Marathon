namespace Marathon
{
    public class TestRunner
    {
        private static void Main(string[] buildArguments)
        {
            FrameworkEngine.EnableExecutorMode(buildArguments);
            UnifiedFunctionalTesting.Close();
            UnifiedFunctionalTesting.Execute();
            UnifiedFunctionalTesting.CleanUp();
        }
    }
}