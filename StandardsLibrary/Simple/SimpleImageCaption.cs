using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleImageCaption
    {
        [DisplayName("Стиль")]
        public SimpleStyle Style { get; set; }
        [DisplayName("Формат нумерации")]
        public string Format { get; set; } = NumberingEnum.Section.ToString();
        [DisplayName("Правила нумерации")]
        public int Rule { get; set; } = 3;

        public enum NumberingEnum
        {
            [Description("Одинокая цифра")] Simple,
            [Description("Цифра с разделом")] Section
        }
    }
}
