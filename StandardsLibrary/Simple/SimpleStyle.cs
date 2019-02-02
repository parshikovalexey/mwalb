using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleStyle
    {
        [DisplayName("Шрифт")] public string Font { get; set; }
        [DisplayName("Размер шрифта")] public int FontSize { get; set; }
        [DisplayName("Межстрочный интервал")] public float LineSpacing { get; set; }
        [DisplayName("Интервал перед абзацем")] public float BeforeSpacing { get; set; }
        [DisplayName("Интервал после абзацем")] public float AfterSpacing { get; set; }
        [DisplayName("Дополнительный отступ к первой строке")] public float FirstLineIndentation { get; set; }
        [DisplayName("Отступ слева")] public float LeftIndentation { get; set; }
        [DisplayName("Отступ справа")] public float RightIndentation { get; set; }
        [DisplayName("Выравнивание")] public string Alignment { get; set; }
        [DisplayName("Поля: левое")] public float MarginLeft { get; set; }
        [DisplayName("Поля: правое")] public float MarginRight { get; set; }
        [DisplayName("Поля: верхнее")] public float MarginTop { get; set; }
        [DisplayName("Поля: нижнее")] public float MarginBottom { get; set; }
        [DisplayName("Цвет текста")] public string Color { get; set; }
        [DisplayName("Жирный")] public bool Bold { get; set; }



    }
}
