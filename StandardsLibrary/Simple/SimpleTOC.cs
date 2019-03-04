using System.ComponentModel;

namespace StandardsLibrary
{
    public class SimpleTOC
    {
        [DisplayName("Заполнитель табуляции")] public string TabLeader { get; set; }
        [DisplayName("Позиция номера страницы")] public float TabPosition { get; set; }
        [DisplayName("Отступ нулевого уровня")] public float LeftIndentation { get; set; }
        [DisplayName("Отступ каждого следующего уровня")] public float LeftNextIndentation { get; set; }
        [DisplayName("Отступ первой строки")] public float FirstLineIndentation { get; set; }
    }
}