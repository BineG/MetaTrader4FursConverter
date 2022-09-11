using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MetaTraderConversionLib
{
    internal class ArchiveParser
    {
        public XmlDocument Parse(string filename)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(ReadFile(filename));

            return doc;
        }

        private Stream ReadFile(string filename)
        {
            return File.OpenRead(filename);
        }
    }
}
