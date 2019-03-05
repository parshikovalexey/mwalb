using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleCaptionPictureTable
    {
        [DisplayName("Формат нумерации")]
        public string NumberingFormat { get; set; } = NumberingEnum.Through.ToString();

        public enum NumberingEnum
        {
            [Description("Сквозная")] Through,
            [Description("Раздел")] Section,
            [Description("Подраздел")] Subsection,
        }
    }
}
