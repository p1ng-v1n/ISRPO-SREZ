using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Srez.Models
{
    internal class Sale
    {
        public DateTime dateSale { get; set; }
        public Client client { get; set; }

        public Telephone[] telephones { get; set; }
    }
}
