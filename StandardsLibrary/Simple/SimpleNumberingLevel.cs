using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleNumberingLevel
    {
        [DisplayName("Формат нумерации")]
        public string NumberingFormat { get; set; } = NumberingEnum.NoFormat.ToString();
        [DisplayName("Шаблон нумерации уровня")]
        public string LevelText { get; set; }
        [DisplayName("Выравнивание уровня")]
        public string LevelJustification { get; set; }

        public enum NumberingEnum
        {
            [Description("Маркер")] Bullet,
            [Description("Десятичные числа")] Decimal,
            [Description("Строчная буква алфавита латиница")] LowerLetter,
            [Description("Римские цифры в нижнем регистре")] LowerRoman,
            [Description("Не менять формат")] NoFormat,
        }
    }
}
