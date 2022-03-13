using MarsFramework.Pages;
using System;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using RelevantCodes.ExtentReports;
using static MarsFramework.Global.GlobalDefinitions;
using MarsFramework.Config;
#nullable disable


namespace MarsFramework.Global
{
    class Base
    {
        #region To access Path from resource file

        public static int Browser = Int32.Parse(MarsResource.Browser);
        //public static string Browser = MarsResource.Browser;
        public static string ExcelPath = MarsResource.ExcelPath;
        public static string ScreenshotPath = MarsResource.ScreenShotPath;
        public static string ReportPath = MarsResource.ReportPath;
        public static string FilePath = MarsResource.FilePath;
        #endregion

        #region reports
        public static ExtentTest test;
        public static ExtentReports extent;

        #endregion

        #region setup and tear down
        [OneTimeSetUp]
        public void Inititalize()
        {
            switch (Browser)
            {
                case 1:
                    GlobalDefinitions.driver = new FirefoxDriver();
                    break;
                case 2:
                    GlobalDefinitions.driver = new ChromeDriver();

                    GlobalDefinitions.driver.Manage().Window.Maximize();
                    break;
            }

            //Populate the excel data
            Thread.Sleep(5000);
            GlobalDefinitions.ExcelLib.PopulateInCollection(Base.ExcelPath, "SignIn");
            GlobalDefinitions.driver.Navigate().GoToUrl(GlobalDefinitions.ExcelLib.ReadData(2, "Url"));

            #region Initialize Reports

            extent = new ExtentReports(ReportPath, false, DisplayOrder.NewestFirst);
            extent.LoadConfig(MarsResource.ReportXMLPath);

            #endregion

            if (MarsResource.IsLogin == "true")
            {
                //Create Extent Report
                test = extent.StartTest("SignIn", "Sample description");
                //SignIn
                SignIn loginobj = new SignIn();
                loginobj.LoginSteps();
            }
            else
            {
                //Create Extent Report
                test = extent.StartTest("Join", "Sample description");
                //Join
                SignUp obj = new SignUp();
                obj.register();
            }
        }

       [OneTimeTearDown]
        //[Test]
        public void TearDown()
        {
            // Screenshot
            String img = SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.driver, "Report");//AddScreenCapture(@"E:\Dropbox\VisualStudio\Projects\Beehive\TestReports\ScreenShots\");
            test.Log(LogStatus.Info, "Image example: " + img);

            // end test. (Reports)
            extent.EndTest(test);

            // calling Flush writes everything to the log file (Reports)
            extent.Flush();

            // Close the driver :)            
            GlobalDefinitions.driver.Close();
            GlobalDefinitions.driver.Quit();
        }
        #endregion
    }
}