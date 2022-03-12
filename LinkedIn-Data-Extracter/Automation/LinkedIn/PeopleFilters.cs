using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;

namespace LinkedIn_Data_Extracter.Automation.LinkedIn
{
    public class PeopleFilters
    {
        private IWebDriver _webDriver;
        public PeopleFilters(IWebDriver webDriver)
        {
            _webDriver = webDriver;
        }

        public void BuildFilterCompany(string company)
        {
            OpenFilter();

            System.Threading.Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath("//*[text()='Add a company']")).Click();

            System.Threading.Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath("//input[@placeholder='Add a company']")).SendKeys(company);

            System.Threading.Thread.Sleep(2000);
            _webDriver.FindElement(By.XPath("//div[contains(@id,'triggered-expanded-')]/div[1]/div[1]")).Click();

            System.Threading.Thread.Sleep(3000);
        }

        public void BuildFilterLocation(string location)
        {
            OpenFilter();

            System.Threading.Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath("//*[text()='Add a location']")).Click();

            System.Threading.Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath("//input[@placeholder='Add a location']")).SendKeys(location);

            System.Threading.Thread.Sleep(2000);
            _webDriver.FindElement(By.XPath("//div[contains(@id,'triggered-expanded-')]/div[1]/div[1]")).Click();

        }

        public void ApplyFilter()
        {
            System.Threading.Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath("//*[text()='Show results']")).Click();

            System.Threading.Thread.Sleep(3000);
        }

        public void OpenFilter()
        {
            IWebElement filterModal = _webDriver.FindElement(By.XPath("//div[contains(@id,'artdeco-modal-outlet')]"));
            if (filterModal.Text.Equals(""))
            {
                System.Threading.Thread.Sleep(5000);
                _webDriver.FindElement(By.XPath("//*[text()='All filters']")).Click();
            }

        }
    }
}
