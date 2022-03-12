using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using LinkedIn_Data_Extracter.ViewModels;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace LinkedIn_Data_Extracter
{
    class Program
    {   
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Console.WriteLine("Welcome to LinkedIn Data Extractor. Broswer will be opened when you select option any option");
            Console.WriteLine("Hang tight during data extract!!!");
            while (true)
            {
                Console.WriteLine("Press 1 for linkedIn. 2 for facebook. 3 for glassdoor");
                string option = Console.ReadLine();
                StartDataExtraction(option);
            }
        }

        private static void StartDataExtraction(string option)
        {
            switch (option)
            {
                case "1":
                    LinkedInAutomation linkedInAutomation = new LinkedInAutomation();
                    linkedInAutomation.ExtractfromLinkedIn();                    
                    break;
                case "2":
                    Console.WriteLine("This option is yet to integrate!!! Contact us for more details");
                    break;
                case "3":
                    Console.WriteLine("This option is yet to integrate!!! Contact us for more details");
                    break;
                default:
                    Console.WriteLine("Please provide a valid option");
                    break;
            }
        }
    }
}
