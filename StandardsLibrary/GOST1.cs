using System;
using System.IO;
using Newtonsoft.Json;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandardsLibrary
{
 public class Gost1:Standards
    {
        public string Font { get; set; } 
        public int FontSize { get; set; } 
        public float LineSpacing { get; set; }
        public string Alignment { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }

        public Gost1(string path)
        {
            Dictionary<string,string> gost = JsonConvert.DeserializeObject<Dictionary<string,string>>(File.ReadAllText(path));
            Font = gost["Font"];
            FontSize = Int32.Parse(gost["FontSize"]);
            LineSpacing = Convert.ToSingle(gost["LineSpacing"], new CultureInfo("en-US"));
            Alignment = gost["Alignment"];
            MarginLeft = Int32.Parse(gost["MarginLeft"]);
            MarginRight = Int32.Parse(gost["MarginRight"]);
            MarginTop = Int32.Parse(gost["MarginTop"]);
            MarginBottom = int.Parse(gost["MarginBottom"]);
        }

        public override int BodyFontSize()
        {
            throw new NotImplementedException();
        }

        public override int HeaderFontSize()
        {
            throw new NotImplementedException();
        }
    }
}
