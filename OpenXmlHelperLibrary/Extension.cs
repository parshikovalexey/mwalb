using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlHelperLibrary
{
    public static class Extension
    {
        public static JustificationValues GetJustificationByString(this string justificationVol, JustificationValues @default = JustificationValues.Both)
        {
            switch (justificationVol.ToLower())
            {
                case "center": return JustificationValues.Center;
                case "left": return JustificationValues.Left;
                case "right": return JustificationValues.Right;
                case "both": return JustificationValues.Both;
            }
            return @default;
        }

        public static LevelJustificationValues GetLevelJustificationByString(this string justificationVol, LevelJustificationValues @default = LevelJustificationValues.Left)
        {
            switch (justificationVol.ToLower())
            {
                case "center": return LevelJustificationValues.Center;
                case "left": return LevelJustificationValues.Left;
                case "right": return LevelJustificationValues.Right;
            }
            return @default;
        }

        public static UnderlineValues GetUnderlineByString(this string underlineVol, UnderlineValues @default = UnderlineValues.None)
        {
            switch (underlineVol.ToLower())
            {
                case "single": return UnderlineValues.Single;
                case "double": return UnderlineValues.Double;
                case "dash": return UnderlineValues.Dash;
                case "thick": return UnderlineValues.Thick;
                case "wave": return UnderlineValues.Wave;
                case "none": return UnderlineValues.None;
                case "words": return UnderlineValues.Words;
            }
            return @default;
        }
        public static NumberFormatValues GetNumberingFormat(this string vol, NumberFormatValues @default = NumberFormatValues.Decimal)
        {

            switch (vol.ToLower().Trim())
            {
                case "bullet": return NumberFormatValues.Bullet;
                case "decimal": return NumberFormatValues.Decimal;
                case "lowerletter": return NumberFormatValues.LowerLetter;
                case "lowerroman": return NumberFormatValues.LowerRoman;
                case "noformat": return NumberFormatValues.None;
                default: return @default;
            }

        }
    }
}
