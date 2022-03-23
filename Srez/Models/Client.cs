using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Srez.Models
{
    internal class Client
    {
        public string lastName { get; set; }
        public string firstName { get; set; }
        public string patronymic { get; set; }
        public string fullname
        {
            get
            {
                return lastName + " " + firstName.Substring(0, 1) + ". " + patronymic.Substring(0, 1) + ".";
            }
        }
    }
}
