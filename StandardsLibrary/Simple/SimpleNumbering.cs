using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleNumbering
    {
       
        [DisplayName("Отступ 0 уровень")] public float LeftIndentation { get; set; }
        [DisplayName("Отступ следующих уровней")] public float LeftNextIndentation { get; set; }
        [DisplayName("Отступ номера")] public float Hanging { get; set; }
        [DisplayName("Стиль Уровень 1")] public SimpleNumberingLevel Level1 { get; set; } 
        [DisplayName("Стиль Уровень 2")] public SimpleNumberingLevel Level2 { get; set; }
        [DisplayName("Стиль Уровень 3")] public SimpleNumberingLevel Level3 { get; set; }
        [DisplayName("Стиль Уровень 4")] public SimpleNumberingLevel Level4 { get; set; }
        [DisplayName("Стиль Уровень 5")] public SimpleNumberingLevel Level5 { get; set; }
        [DisplayName("Стиль Уровень 6")] public SimpleNumberingLevel Level6 { get; set; }
        [DisplayName("Стиль Уровень 7")] public SimpleNumberingLevel Level7 { get; set; }
        [DisplayName("Стиль Уровень 8")] public SimpleNumberingLevel Level8 { get; set; }
        [DisplayName("Стиль Уровень 9")] public SimpleNumberingLevel Level9 { get; set; }


     
    }
}
