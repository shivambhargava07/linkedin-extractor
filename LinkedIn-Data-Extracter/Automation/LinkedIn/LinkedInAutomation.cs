using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using LinkedIn_Data_Extracter.Automation.LinkedIn;
using LinkedIn_Data_Extracter.ViewModels;
using Microsoft.VisualBasic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;

namespace LinkedIn_Data_Extracter
{
    public class LinkedInAutomation
    {
        private readonly IWebDriver _webDriver;
        private readonly string _userName;
        private readonly string _password;
        private readonly string _filePath;
        private readonly bool _isWebsiteRequired;
        private readonly bool _isSalesNavigatorEnabled;
        private readonly string _loginURL;
        private readonly string _companyURL;
        private readonly string _peopleURL;
        private readonly string _salesNavigtorURL;
        private readonly bool _pagingEnabled;
        private readonly int _pageSize;
        private readonly PeopleFilters _peopleFilters;
        private readonly SalesNavigationFilters _salesNavigatorFilters;

        public LinkedInAutomation()
        {
            _webDriver = LaunchBrowser();
            _userName = ConfigurationManager.AppSettings.Get("userName");
            _password = ConfigurationManager.AppSettings.Get("password");
            _filePath = ConfigurationManager.AppSettings.Get("filePath");
            _isWebsiteRequired = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("isWebsiteRequired"));
            _isSalesNavigatorEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("isSalesNavigatorEnabled"));
            _loginURL = ConfigurationManager.AppSettings.Get("linKedInLoginURL");
            _companyURL = ConfigurationManager.AppSettings.Get("linKedInCompanyURL");
            _peopleURL = ConfigurationManager.AppSettings.Get("linKedInPeopleURL");
            _salesNavigtorURL = ConfigurationManager.AppSettings.Get("linKedInSalesNavigatorURL");
            _peopleFilters = new PeopleFilters(_webDriver);
            _salesNavigatorFilters = new SalesNavigationFilters(_webDriver);
            _pagingEnabled = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("pagingEnabled"));
            _pageSize = Convert.ToInt32(ConfigurationManager.AppSettings.Get("pageSize"));
        }

        #region Public Constructors

        public void ExtractfromLinkedIn()
        {
            try
            {
                List<LinkedInViewModel> linkedInDetailedList = GetLinkedInDetails();

                if (linkedInDetailedList.Count > 0)
                {
                    //login
                    Login();

                    if (_isSalesNavigatorEnabled)
                    {
                        linkedInDetailedList = ExtractSalesNavigatorPeopleProfiles(linkedInDetailedList);
                    }
                    else
                    {
                        linkedInDetailedList = ExtractPeopleProfiles(linkedInDetailedList);
                    }

                    if (_isWebsiteRequired)
                    {
                        linkedInDetailedList = ExtractCompanyWebsite(linkedInDetailedList);
                    }

                    DataTable dt = BuildTableforExcelExport(linkedInDetailedList);

                    //Name of File  
                    string fileName = "OutputFile.xlsx";
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        string saveTo = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        //Add DataTable in worksheet  
                        wb.Worksheets.Add(dt, "Data");
                        wb.SaveAs(string.Concat(saveTo, @"\", fileName));
                        Console.WriteLine("Excel Exported Successfully at :- {0}", saveTo);
                    }
                }
                else
                {
                    Console.WriteLine("No Data found from given excel file");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Application Crashed with an exception", ex);
            }
            finally
            {
                Console.WriteLine("Data Extraction completed successfully");
                _webDriver.Quit();
            }
        }

        public void Login()
        {
            try
            {
                _webDriver.Url = _loginURL;
                _webDriver.FindElement(By.Id("username")).SendKeys(_userName);
                _webDriver.FindElement(By.Id("password")).SendKeys(_password);

                // Click on the login button
                ClickAndWaitForPageToLoad(_webDriver, By.XPath("//*[@id='organic-div']/form/div[3]/button"));
                Console.WriteLine("User Logged In Successfully!!!");

                //Skip Optional Screen
                //SkipOptionalScreenOnNewMachine();
            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Element with locator was not found {0}", ex);
                throw;
            }
            catch (Exception)
            {
                Console.WriteLine("An excpetion occured while loggin-in");
                throw;
            }
        }

        public List<LinkedInViewModel> ExtractCompanyWebsite(List<LinkedInViewModel> linkedInDetailedList)
        {
            List<LinkedInViewModel> buildCompanyDomainResult = new List<LinkedInViewModel>();
            foreach (LinkedInViewModel linkedInDetails in linkedInDetailedList)
            {
                if (linkedInDetails.EmployeeProfiles == null)
                {
                    continue;
                }
                foreach (ProfileViewModel people in linkedInDetails.EmployeeProfiles)
                {
                    try
                    {
                        _webDriver.Url = _companyURL + "/?keywords=" + people.EmployeeCompany;

                        _peopleFilters.BuildFilterLocation(people.EmployeeLocation);

                        _peopleFilters.ApplyFilter();

                        try
                        {
                            //core logic of this method lies here
                            //Go inside company
                            System.Threading.Thread.Sleep(2000);
                            var comapnyAnchorLink = _webDriver.FindElements(By.ClassName("app-aware-link"));
                            if (comapnyAnchorLink.Count >= 2)
                            {
                                comapnyAnchorLink[1].Click();
                            }

                            //fetch website
                            System.Threading.Thread.Sleep(2000);
                            _webDriver.FindElement(By.LinkText("About")).Click();

                            System.Threading.Thread.Sleep(2000);
                            var websiteLink = _webDriver.FindElement(By.CssSelector(".link-without-visited-state.ember-view"));
                            string webSite = websiteLink.GetAttribute("href");

                            Console.WriteLine("{0} website is :- {1}", webSite);
                            //#TODO
                            people.EmployeeCompanyWebsite = webSite;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("{0} does not have website on linkedin", people.EmployeeCompany);
                        }

                        System.Threading.Thread.Sleep(2000);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error occured while extracting comapny website for {0} with exception details :- {1}", people.EmployeeCompany, ex);
                    }
                }
            }
            return linkedInDetailedList;
        }

        public List<LinkedInViewModel> ExtractPeopleProfiles(List<LinkedInViewModel> linkedInDetailedList)
        {
            foreach (LinkedInViewModel profiledetails in linkedInDetailedList)
            {
                try
                {
                    _webDriver.Url = _peopleURL;

                    _peopleFilters.BuildFilterCompany(profiledetails.CompanyName);

                    _peopleFilters.BuildFilterLocation(profiledetails.LocationName);

                    _peopleFilters.ApplyFilter();

                    System.Threading.Thread.Sleep(2000);
                    try
                    {
                        profiledetails.EmployeeProfiles = ExtractEmployeeProfile();
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("employee profiles not found: {0}", profiledetails.CompanyName);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error occured while extracting", ex);
                }
            }
            return linkedInDetailedList;
        }

        public List<LinkedInViewModel> ExtractSalesNavigatorPeopleProfiles(List<LinkedInViewModel> linkedInDetailedList)
        {
            foreach (LinkedInViewModel profiledetails in linkedInDetailedList)
            {
                try
                {
                    _webDriver.Url = _salesNavigtorURL;

                    Thread.Sleep(3000);
                    _salesNavigatorFilters.ExpandSearch();
                    //keywords filter
                    if (!string.IsNullOrEmpty(profiledetails.KeyWords))
                    {
                        _salesNavigatorFilters.BuildFilterKeyWords(profiledetails.KeyWords);
                    }
                    
                    //location/geography filter
                    if (!string.IsNullOrEmpty(profiledetails.LocationName))
                    {
                        _salesNavigatorFilters.BuildFilterLocation(profiledetails.LocationName);
                    }

                    //company filter
                    if (!string.IsNullOrEmpty(profiledetails.CompanyName))
                    {
                        _salesNavigatorFilters.BuildFilterCompany(profiledetails.CompanyName);
                    }

                    //industry filter
                    if (!string.IsNullOrEmpty(profiledetails.Industries))
                    {
                        _salesNavigatorFilters.BuildFilterIndustries(profiledetails.Industries);
                    }

                    //function filter
                    if (!string.IsNullOrEmpty(profiledetails.Functions))
                    {
                        _salesNavigatorFilters.BuildFilterFunctions(profiledetails.Functions);
                    }

                    //title filter
                    if (!string.IsNullOrEmpty(profiledetails.Titles))
                    {
                        _salesNavigatorFilters.BuildFilterTitles(profiledetails.Titles);
                    }

                    try
                    {
                        profiledetails.EmployeeProfiles = ExtractEmployeeProfile();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("employee profiles was not found for {0} with ex : {1}", profiledetails.LocationName, ex);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error occured while extracting", ex);
                }
            }
            return linkedInDetailedList;
        }

        public List<LinkedInViewModel> GetLinkedInDetails()
        {
            List<LinkedInViewModel> linkedInDetailsList = new List<LinkedInViewModel>();
            using (var stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    foreach (DataTable dt in result.Tables)
                    {
                        if (dt.TableName.Equals("Sheet1"))
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                linkedInDetailsList.Add(new LinkedInViewModel()
                                {
                                    CompanyName = dr["CompanyName"].ToString(),
                                    LocationName = dr["LocationName"].ToString(),
                                    KeyWords = dr["KeyWords"].ToString(),
                                    Functions = dr["Functions"].ToString(),
                                    Titles = dr["Titles"].ToString(),
                                    Industries = dr["Industries"].ToString()
                                });
                            }
                        }
                    }
                }
            }
            return linkedInDetailsList;
        }

        #endregion

        #region Private Methods

        private IWebDriver LaunchBrowser()
        {
            var options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArgument("--disable-notifications");

            var driver = new ChromeDriver(Environment.CurrentDirectory, options);
            return driver;
        }

        private void SkipOptionalScreenOnNewMachine()
        {
            // At this point, Facebook will launch a post-login "wizard" that will 
            // keep asking unknown amount of questions (it thinks it's the first time 
            // you logged in using this computer). We'll just click on the "continue" 
            // button until they give up and redirect us to our "wall".
            while (_webDriver.FindElement(By.Id("skip")) != null)
            {
                // Clicking "continue" until we're done
                ClickAndWaitForPageToLoad(_webDriver, By.Id("checkpointSubmitButton"));
            }
        }

        private void ClickAndWaitForPageToLoad(IWebDriver driver, By elementLocator, int timeout = 100)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
            var elements = driver.FindElements(elementLocator);
            if (elements.Count == 0)
            {
                throw new NoSuchElementException(
                    "No elements " + elementLocator + " ClickAndWaitForPageToLoad");
            }
            var element = elements.FirstOrDefault(e => e.Displayed);
            element.Click();
            wait.Until(ExpectedConditions.StalenessOf(element));
        }

        private List<ProfileViewModel> ExtractEmployeeProfile()
        {
            List<ProfileViewModel> employeeProfiles = new List<ProfileViewModel>();
            bool IsNextButton = true;
            System.Threading.Thread.Sleep(2000);
            try
            {
                while (IsNextButton)
                {
                    System.Threading.Thread.Sleep(2000);

                    MovePageUpandDown(_webDriver);

                    System.Threading.Thread.Sleep(2000);
                    int pagePageCount = Convert.ToInt32(_webDriver.FindElements(By.XPath("//ol[contains(@class,'artdeco-list')]/li")).Count);

                    for (int i = 1; i <= pagePageCount; i++)
                    {

                        System.Threading.Thread.Sleep(1000);
                        ProfileViewModel profile = new ProfileViewModel();

                        try
                        {
                            //employeename                
                            profile.EmployeeName = _webDriver.FindElement(By.XPath("//ol[contains(@class,'artdeco-list')]/li[" + i + "]/div/div/div[2]/div/div/div/div[2]/div/div/a")).Text;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Employee Name Profile not found or visible to selenium for row {0} with execption :- {1}", i, ex);
                        }

                        try
                        {
                            profile.EmployeeTitle = _webDriver.FindElement(By.XPath("//ol[contains(@class,'artdeco-list')]/li[" + i + "]/div/div/div[2]/div/div/div/div[2]/div[2]/span[1]")).Text;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Employee Title Profile not found or visible to selenium for row {0} with execption :- {1}", i, ex);
                        }

                        try
                        {
                            string fullTitleWithCompany = _webDriver.FindElement(By.XPath("//ol[contains(@class,'artdeco-list')]/li[" + i + "]/div/div/div[2]/div/div/div/div[2]/div[2]")).Text;
                            profile.EmployeeCompany = fullTitleWithCompany.Replace(profile.EmployeeTitle,"").Trim();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Employee Company Profile not found or visible to selenium for row {0} with execption :- {1}", i, ex);
                        }

                        try
                        {
                            //employeelocation                
                            profile.EmployeeLocation = _webDriver.FindElement(By.XPath("//ol[contains(@class,'artdeco-list')]/li[" + i + "]/div/div/div[2]/div/div/div/div[2]/div[3]/span")).Text;
                            employeeProfiles.Add(profile);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Employee Location Profile not found or visible to selenium for row {0} with execption :- {1}", i, ex);
                        }

                    }

                    if (_pagingEnabled)
                    {
                        IWebElement currentPagingElement = _webDriver.FindElement(By.XPath("//li[contains(@class,'artdeco-pagination__indicator artdeco-pagination__indicator--number active selected')]/button/span[1]"));
                        int currentPageValue = Convert.ToInt32(currentPagingElement.Text);
                        Console.WriteLine(currentPageValue);
                        if (currentPageValue > _pageSize)
                        {
                            IsNextButton = false;
                        }
                        else
                        {
                            //check for paging
                            IWebElement nextbutton = _webDriver.FindElement(By.XPath("//button[contains(@class,'artdeco-pagination__button--next')]"));

                            if (nextbutton.Displayed)
                            {
                                IsNextButton = true;
				nextbutton.SendKeys(Keys.Enter);
                                Thread.Sleep(4000);
                            }
                            else
                            {
                                IsNextButton = false;
                            }
                        }
                    }
                    else
                    {
                        IsNextButton = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled Exception {0}", ex);
            }
            return employeeProfiles;
        }

        private void MovePageUpandDown(IWebDriver webDriver)
        {
            webDriver.FindElement(By.TagName("body")).SendKeys(Keys.Home);
            System.Threading.Thread.Sleep(1000);

            IWebElement ele = webDriver.FindElement(By.XPath("//div[contains(@class,'_bulk-save-checkbox-spacing')]/input"));
            
            webDriver.ExecuteJavaScript("arguments[0].click()", ele);
            Thread.Sleep(2000);

            _webDriver.FindElement(By.XPath("//div[@id='search-results-container']")).Click();
            Thread.Sleep(2000);
            for (int x = 0; x < 11; x++)
            {
                System.Threading.Thread.Sleep(500);
                webDriver.FindElement(By.TagName("body")).SendKeys(Keys.PageDown);
            }
        }

        public DataTable BuildTableforExcelExport(List<LinkedInViewModel> linkedInDetailedList)
        {
            DataTable table = new DataTable();
            table.Columns.Add("CompanyName", typeof(string));
            table.Columns.Add("LocationName", typeof(string));
            table.Columns.Add("KeyWords", typeof(string));
            table.Columns.Add("Functions", typeof(string));
            table.Columns.Add("Titles", typeof(string));
            table.Columns.Add("Industries", typeof(string));
            table.Columns.Add("EmployeeName", typeof(string));
            table.Columns.Add("EmployeeTitle", typeof(string));
            table.Columns.Add("EmployeeLocation", typeof(string));
            table.Columns.Add("EmployeeCompany", typeof(string));
            table.Columns.Add("EmployeeCompanyWebsite", typeof(string));
            table.Columns.Add("EmployeeCompanyEmailId", typeof(string));

            foreach (LinkedInViewModel linkedInDetailed in linkedInDetailedList)
            {
                if (linkedInDetailed.EmployeeProfiles == null)
                {
                    table.Rows.Add(linkedInDetailed.CompanyName, linkedInDetailed.LocationName, linkedInDetailed.KeyWords, linkedInDetailed.Functions, linkedInDetailed.Titles, linkedInDetailed.Industries);
                }
                else
                {
                    foreach (ProfileViewModel profileDetails in linkedInDetailed.EmployeeProfiles)
                    {
                        table.Rows.Add(linkedInDetailed.CompanyName, linkedInDetailed.LocationName, linkedInDetailed.KeyWords, linkedInDetailed.Functions, linkedInDetailed.Titles, linkedInDetailed.Industries, profileDetails.EmployeeName, profileDetails.EmployeeTitle, profileDetails.EmployeeLocation, profileDetails.EmployeeCompany, profileDetails.EmployeeCompanyWebsite, profileDetails.EmployeeCompanyEmailId);
                    }
                }

            }
            return table;
        }

        #endregion

    }
}
