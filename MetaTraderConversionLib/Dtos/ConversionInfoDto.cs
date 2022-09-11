using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaTraderConversionLib.Dtos
{
    public class ConversionInfoDto
    {
        public PersonInfoDto PersonInfo { get; set; } = new PersonInfoDto();

        public DocumentInfoDto DocumentInfo { get; set; } = new DocumentInfoDto();
    }
}
