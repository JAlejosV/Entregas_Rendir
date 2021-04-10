using System;
using System.Collections.Generic;
using System.Text;

namespace EntregasRendir.Context.Models
{
    class User
    {
        public string Id { get; set; }
        public string FirstName { get; set; }
        public string PreferredName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string FullName { get { return $"{FirstName} {LastName}"; } set { } }
        public string Suffix { get; set; }
        public string LdapName { get; set; }
        public string Email { get; set; }
        //public string EmployeeNumber { get; set; }
        public string HomePhoneNumber { get; set; }
        public string CellPhoneNumber { get; set; }
        public string Birthday { get; set; }

        private string _EmployeeNumber;
        //Sirve para en caso vengan valores nulos, se asigne valor en blanco.
        public string EmployeeNumber
        {
            get { return _EmployeeNumber ?? ""; }
            set { _EmployeeNumber = value; }
        }

    }
}
