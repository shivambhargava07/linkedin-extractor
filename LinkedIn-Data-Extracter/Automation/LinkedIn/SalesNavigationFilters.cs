﻿using OpenQA.Selenium;
using System;
using System.Threading;

namespace LinkedIn_Data_Extracter.Automation.LinkedIn
{
    public class SalesNavigationFilters
    {
        private IWebDriver _webDriver;
        public SalesNavigationFilters(IWebDriver webDriver)
        {
            _webDriver = webDriver;
        }

        public void ExpandSearch()
        {
            try
            {
                Thread.Sleep(1000);
                _webDriver.FindElement(By.XPath("//button[contains(@data-test-search-toggle-panel,'expand')]")).Click();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Unable to toggle side filter", ex);
            }
        }

        public void BuildFilterKeyWords(string keywords)
        {
            try
            {
                Thread.Sleep(1000);
                _webDriver.FindElement(By.XPath("//input[@placeholder='Search keywords']")).Click();

                _webDriver.FindElement(By.XPath("//input[@placeholder='Search keywords']")).SendKeys(keywords.Trim());
                Thread.Sleep(1000);
                _webDriver.FindElement(By.XPath("//input[@placeholder='Search keywords']")).SendKeys(Keys.Enter);
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to find keywords {0) with exception :- {1}", keywords, ex);
            }
        }

        public void BuildFilterCompany(string company)
        {
            try
            {
                ExpandBlockFilter("Current Company");
                PasteIntoFilter("Add current companies", company.Trim());
                ClickFilterList();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Unable to find comapny {0} with exception :-", company, ex);
            }
        }

        public void BuildFilterLocation(string location)
        {
            try
            {
                ExpandBlockFilter("Geography");
                PasteIntoFilter("Add locations", location.Trim());
                ClickFilterList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to filter location {0} with exception :- {1}", location, ex);
            }
        }

        public void BuildFilterFunctions(string functions)
        {
            try
            {
                ExpandBlockFilter("Function");
                string[] functionsArray = functions.Split(",");
                foreach (string function in functionsArray)
                {
                    PasteIntoFilter("Add functions", function.Trim());
                    ClickFilterList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to find function {0} with exception :- {1}", functions, ex);
            }
        }

        public void BuildFilterTitles(string titles)
        {
            try
            {
                ExpandBlockFilter("Job title");

                string[] titleArray = titles.Split(",");
                foreach (string title in titleArray)
                {
                    PasteIntoFilter("Add titles", title.Trim());
                    ClickFilterList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to find title {0} with exception :- {1}", titles, ex);
            }
        }

        public void BuildFilterIndustries(string industries)
        {
            try
            {
                ExpandBlockFilter("Industry");

                string[] industryArray = industries.Split(",");
                foreach (string industry in industryArray)
                {
                    PasteIntoFilter("Add industries", industry.Trim());
                    ClickFilterList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to find industry {0} with exception :- {1}", industries, ex);
            }
        }

        private void ExpandBlockFilter(string filterName)
        {
            Thread.Sleep(1000);
            _webDriver.FindElement(By.XPath($"//legend[text()='{filterName}']//following-sibling::button")).Click();
        }

        private void PasteIntoFilter(string placeholderName, string filterValue)
        {
            _webDriver.FindElement(By.XPath($"//input[@placeholder='{placeholderName}']")).SendKeys(filterValue);
        }

        private void ClickFilterList()
        {
            Thread.Sleep(3000);
            _webDriver.FindElement(By.XPath($"//ul[contains(@data-test-search-filter-typeahead,'suggestions-list')]//li[1]/div/button[1]")).Click();
        }
    }
}