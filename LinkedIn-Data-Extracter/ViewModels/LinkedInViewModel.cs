using System;
using System.Collections.Generic;
using System.Text;

namespace LinkedIn_Data_Extracter.ViewModels
{
    public class LinkedInViewModel
    {
        public string CompanyName { get; set; }

        public string LocationName { get; set; }

        public string KeyWords { get; set; }

        public string Functions { get; set; }

        public string Titles { get; set; }

        public string Industries { get; set; }
        
        public List<ProfileViewModel> EmployeeProfiles { get; set; }
    }
}
