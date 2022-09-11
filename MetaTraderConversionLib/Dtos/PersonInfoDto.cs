using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaTraderConversionLib.Dtos
{
    public class PersonInfoDto
    {
        public string NameSurname { get; set; } = "";
        public string Address1 { get; set; } = "";
        public string City { get; set; } = "";
        public string PostNumber { get; set; } = "";
        public string TaxNumber { get; set; } = "";
        public string Telephone { get; set; } = "";
        public string Email { get; set; } = "";
        public DateTime Birthdate { get; set; }
    }
}
