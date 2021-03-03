﻿using System;
using System.Threading.Tasks;
using System.Threading;
using System.Management.Automation;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Chrome;
using CustomExceptions;


namespace SolarWinds_Searcher_Gui
{
    public class SearcherThread : IDisposable
    {
        private readonly ChromeDriver chrome;
        private ExcelInteraction handle;
        private int attributeIndex;
        private int col;
        private int min;
        private int max;

        public SearcherThread(ExcelInteraction handle, int col, int min, int max, int attribute)
        {
            attributeIndex = attribute;
            this.col = col;
            this.min = min;
            this.max = max;
            this.handle = handle;
            Console.WriteLine("Here");
            try
            {
                ChromeOptions options = new ChromeOptions();
                options.AddArgument("headless");
                var chromeDriverService = ChromeDriverService.CreateDefaultService("C:\\Program Files\\ChromeDriver");
                chromeDriverService.HideCommandPromptWindow = true;
                chrome = new ChromeDriver(chromeDriverService, options);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.StackTrace);
                throw new ChromeDriverStartUpException(e.Message);
            }
        }

        private void Search(int row)
        {
            
            string valueToSearch = handle.GetNext(row, col);
            if (valueToSearch != null)
            {
                string textboxName = "ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$MainContentPlaceHolder$ResourceHostControl1$resContainer$rptContainers$ctl00$rptColumn1$ctl00$ctl01$Wrapper$txtSearchString";
                string attributeDropdown = "ctl00$ctl00$ctl00$BodyContent$ContentPlaceHolder1$MainContentPlaceHolder$ResourceHostControl1$resContainer$rptContainers$ctl00$rptColumn1$ctl00$ctl01$Wrapper$lbxNodeProperty";
                string searchBtnId = "ctl00_ctl00_ctl00_BodyContent_ContentPlaceHolder1_MainContentPlaceHolder_ResourceHostControl1_resContainer_rptContainers_ctl00_rptColumn1_ctl00_ctl01_Wrapper_btnSearch";

                chrome.Navigate().GoToUrl("https://solarwindscs.dell.com/Orion/SummaryView.aspx?ViewID=1");

                while (IsElementPresent(textboxName, false, true) == false)
                {
                    Thread.Sleep(25);
                }
                try
                {
                    OpenQA.Selenium.IWebElement searchBox = chrome.FindElementByName(textboxName);
                    OpenQA.Selenium.IWebElement dropBox = chrome.FindElementByName(attributeDropdown);
                    OpenQA.Selenium.IWebElement searchBtn = chrome.FindElementById(searchBtnId);
                    OpenQA.Selenium.Support.UI.SelectElement select = new OpenQA.Selenium.Support.UI.SelectElement(dropBox);
                    //Console.WriteLine("HHHHHHH");
                    searchBox.SendKeys(valueToSearch);
                    //Console.WriteLine("HHHHHHH");
                    select.SelectByIndex(attributeIndex);
                   // Console.WriteLine("uyguygufuov");
                    searchBtn.Click();
                   // Console.WriteLine("HHHHHHH");

                    while (IsElementPresent("StatusMessage", true, false) == false)
                    {
                        Thread.Sleep(25);
                    }

                    string result = chrome.FindElementByClassName("StatusMessage").Text;
                    if (result.Contains("Nodes with ") && result.Contains(" similar to "))
                    {
                        ReadOnlyCollection<OpenQA.Selenium.IWebElement> amount = chrome.FindElementsByClassName("StatusIcon");
                        //Console.WriteLine(amount.Count);
                        handle.AddResult(row, valueToSearch, "Y", (amount.Count - 1).ToString());
                    }
                    else
                    {
                        handle.AddResult(row, valueToSearch, "N", "0");
                    }
                }
                catch (Exception e)
                {
                    throw new WebSearchException(Thread.CurrentThread.ManagedThreadId.ToString());
                }
            }
        }

        public int SearchWrapper()
        {
            try
            {
                for (int i = min; i < max; i++)
                {
                    Search(i);
                }
            }
            catch (Exception e)
            {
                throw new SearcherThreadException(e.Message, e);
            }
            finally
            {
                //handle.DeDupe();
                chrome.Quit();
            }
            return 1;
        }


        private bool IsElementPresent(string val, bool byClass, bool byName)
        {
            try
            {
                if(byClass)
                {
                    var temp = chrome.FindElementByClassName(val); 
                }
                else if(byName)
                {
                    var temp = chrome.FindElementByName(val);
                }
                else
                {
                    var temp = chrome.FindElementById(val);
                }
                return true;
            }
            catch(Exception e)
            {
                return false;
            }
        }

        public void Dispose()
        {
            if(chrome != null)
            {
                chrome.Quit();
            }

        }


    }

}
