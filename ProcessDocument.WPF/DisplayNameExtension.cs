using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace ProcessDocument.WPF
{
    public class DisplayNameExtension : MarkupExtension
    {
        public Type Type { get; set; }

        public string PropertyName { get; set; }

        public DisplayNameExtension() { }
        public DisplayNameExtension(string propertyName)
        {
            PropertyName = propertyName;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            try
            {
                // (This code has zero tolerance)
                var prop = Type.GetProperty(PropertyName);
                var attributes = prop?.GetCustomAttributes(typeof(DisplayNameAttribute), false);
                var result = (attributes[0] as DisplayNameAttribute).DisplayName;
                return result;
            }
            catch (Exception e)
            {
                LoggerLibrary.Logger.Write().Error(e);
                return string.Empty;
            }
          
        }
    }
}
