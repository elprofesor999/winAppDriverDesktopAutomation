using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Diagnostics;


namespace UnitTestProject2
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            String wadPath = "C:\\Program Files (x86)\\Windows Application Driver\\WinAppDriver.exe";
            ProcessStartInfo startInfo = new ProcessStartInfo(wadPath);
            Process wadProcess = Process.Start(startInfo);
            String notepadPath = "C:\\windows\\system32\\notepad.exe";
            WindowsDriver<WindowsElement> notepadSession;
            AppiumOptions desiredCapabilities = new AppiumOptions();
            desiredCapabilities
                .AddAdditionalCapability("app", notepadPath);
            notepadSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723/"), desiredCapabilities);
            notepadSession.FindElementByName("Edit").Click();
            notepadSession.FindElementByAccessibilityId("26").Click();
            wadProcess.Kill();
            
           
            
        }
    }
}
