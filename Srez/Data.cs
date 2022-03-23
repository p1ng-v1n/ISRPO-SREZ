﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Srez
{
    internal class Data
    {
        public class Sale
        {
            public DateTime dateSale { get; set; }
            public Client client { get; set; }

            public Telephone[] telephones { get; set; }
        }

        public class Client
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

        public class Telephone
        {
            public int articul { get; set; }
            public string nameTelephone { get; set; }
            public string category { get; set; }
            public float cost { get; set; }
            public int count { get; set; }
            public string manufacturer { get; set; }
        }
    }
}
