using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
 public   static class CommonHelpers
    {
        /// <summary>
        /// Возвращает значение атриту displayName у Enum
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string DisplayName(this Enum item)
        {
            var type = item.GetType();
            var member = type.GetMember(item.ToString());
            DisplayAttribute displayName = (DisplayAttribute)member[0].GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();
            return displayName != null ? displayName.Name : item.ToString();
        }

    }
}
