using System;

namespace HelperLibrary
{
    public static class SafeConvertExtensions
    {
        public static String nvl(this string s)
        {
            return s ?? "";
        }

        public static int nvl(this int? s)
        {
            return s ?? 0;
        }

        public static bool nvl(this bool? s)
        {
            return s ?? false;
        }

        public static double nvl(this double? s)
        {
            return s ?? 0;
        }

        public static float nvl(this float? s)
        {
            return s ?? 0;
        }


        /// <summary>
        /// Возращает стандартное время 
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static DateTime nvl(this DateTime? s)
        {
            return s ?? default(DateTime);
        }

        /// <summary>
        /// Возвращает текущее время
        /// </summary>
        /// <param name="s"></param>
        /// <returns>Возвращает текущее время</returns>
        public static DateTime nvl2(this DateTime? s)
        {
            return s ?? DateTime.Now;
        }

        /// <summary>
        /// Безопасное конвертирование из double в int 
        /// с указанием значения по умолчанию если конвертирование не удалось
        /// </summary>
        /// <param name="item"></param>
        /// <param name="default"></param>
        /// <returns>в случае успеха вернет преобразованное число, в другом случае возвращает @default</returns>
        public static int SafeToInt(this double item, int @default = 0)
        {
            int i = @default;
            Int32.TryParse(item.ToString(), out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование из double в int 
        /// с указанием значения по умолчанию если конвертирование не удалось
        /// </summary>
        /// <param name="item"></param>
        /// <param name="default"></param>
        /// <returns>в случае успеха вернет преобразованное число, в другом случае возвращает @default</returns>
        public static int SafeToInt(this int? item, int @default)
        {
            if (item == null)
                return @default;
            int i = @default;
            Int32.TryParse(item.ToString(), out i);
            return i;
        }

        public static DateTime SafeToDateTime(this DateTime? item, DateTime @default = default(DateTime))
        {
            if (item == null)
                return @default;
            DateTime i = @default;
            DateTime.TryParse(item.ToString(), out i);
            return i;
        }

        public static bool SafeToBoolean(this int? item, bool @default = false)
        {
            if (item == null)
                return @default;
            int i = 0;
            Int32.TryParse(item.ToString(), out i);
            return i != 0;
            //   return i;
        }

        public static bool SafeToBoolean(this int item, bool @default = false)
        {
            int i = 0;
            Int32.TryParse(item.ToString(), out i);
            return i != 0;
            //   return i;
        }

        public static uint SafeToUint(this int? item, uint @default = 0)
        {
            uint i = @default;
            if (item != null) uint.TryParse(item.ToString(), out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование из decimal в int 
        /// с указанием значения по умолчанию если конвертирование не удалось
        /// </summary>
        /// <param name="item"></param>
        /// <param name="default"></param>
        /// <returns>в случае успеха вернет преобразованное число, в другом случае возвращает @default</returns>
        public static int SafeToInt(this decimal item, int @default = 0)
        {
            var i = @default;
            int.TryParse(item.ToString(), out i);
            return i;
        }


        /// <summary>
        /// Безопасное конвертирование string в int
        /// </summary>
        /// <param name="str"></param>
        /// <param name="default"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает @default</returns>
        public static int SafeToInt(this string str, int @default = 0)
        {
            int i = @default;
            if (!String.IsNullOrEmpty(str))
                Int32.TryParse(str, out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование string в Decimal
        /// </summary>
        /// <param name="str"></param>
        /// <param name="default"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает 0</returns>
        public static decimal SafeToDecimal(this string str, decimal @default = 0)
        {
            decimal i = @default;
            if (!String.IsNullOrEmpty(str))
                Decimal.TryParse(str, out i);
            return i;
        }

        /// <summary>
        ///     Парсинг даты из string
        /// </summary>
        /// <param name="dateTime">дата/время string</param>
        /// <param name="default"></param>
        /// <returns>Возвращает дату/время в случае успеха или default(DateTime) если распарсить дату/время не удалось</returns>
        public static DateTime SafeToDateTime(this string dateTime, DateTime @default = default(DateTime))
        {
            var dt = @default;
            var canParseDt = DateTime.TryParse(dateTime, out dt);
            return dt;
        }

        /// <summary>
        /// Безопасное конвертирование string в Boolean
        /// </summary>
        /// <param name="str"></param>
        /// <returns> в случае успеха вернет преобразованное в логическое значение, в другом случае возвращает false</returns>
        public static bool SafeToBoolean(this string str, bool @default = false)
        {
            bool i = @default;
            if (!String.IsNullOrEmpty(str))
                Boolean.TryParse(str, out i);
            return i;
        }
        /// <summary>
        /// Безопасное конвертирование string в double
        /// </summary>
        /// <param name="str"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает 0</returns>
        public static double SafeToDouble(this string str, double @default = 0)
        {
            double i = @default;
            if (!String.IsNullOrEmpty(str))
                Double.TryParse(str, out i);
            return i;
        }

        public static float SafeToFloat(this string str, float @default = 0)
        {
            float i = @default;
            if (!String.IsNullOrEmpty(str))
                float.TryParse(str, out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование double в double
        /// </summary>
        /// <param name="str"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает 0</returns>
        public static double SafeToDouble(this double? str, double @default = 0)
        {
            double i = @default;
            if (str != null)
                Double.TryParse(str.ToString(), out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование float в double
        /// </summary>
        /// <param name="str"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает 0</returns>
        public static double SafeToDouble(this float str, double @default = 0)
        {
            double i = @default;
            if (!String.IsNullOrEmpty(str.ToString()))
                Double.TryParse(str.ToString(), out i);
            return i;
        }

        /// <summary>
        /// Безопасное конвертирование float? в double
        /// </summary>
        /// <param name="str"></param>
        /// <returns> в случае успеха вернет преобразованное число, в другом случае возвращает 0</returns>
        public static double SafeToDouble(this float? str, double @default = 0)
        {
            if (str == null)
                return @default;
            double i = @default;
            if (!String.IsNullOrEmpty(str.ToString()))
                Double.TryParse(str.ToString(), out i);
            return i;
        }

        public static int ToMargins(this float i)
        {
            return (int)(i * ConstVol.Margin.Multiplier);
        }
    }
}
