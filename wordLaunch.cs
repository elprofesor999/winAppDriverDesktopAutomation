using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace UnitTestProject1
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

            String wordPath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.exe";
            WindowsDriver<WindowsElement> wordSession;
            AppiumOptions desiredCapabilities = new AppiumOptions();
            desiredCapabilities
                .AddAdditionalCapability("app", wordPath);
            wordSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723/"), desiredCapabilities);
            //var currentWindowHandle = wordSession.CurrentWindowHandle;
            ////wordSession.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            //var allWindowHandles = wordSession.WindowHandles;
            //wordSession.SwitchTo().Window(allWindowHandles[1]);
            wordSession.FindElementByName("Blank document").Click();
            //Thread.Sleep(3000);
            //wadProcess.Kill();
        }
       
    }
}
