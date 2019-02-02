using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleNumbering
    {
        [DisplayName("Шрифт")] public string Font { get; set; }
        [DisplayName("Отступ 1 уровень")] public float LeftIndentation { get; set; }
        [DisplayName("Отступ следующих уровней")] public float LeftNextIndentation { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level1 { get; set; } 
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level2 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level3 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level4 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level5 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level6 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level7 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level8 { get; set; }
        [DisplayName("Уровень 1")] public SimpleNumberingLevel Level9 { get; set; }


     
    }
}
