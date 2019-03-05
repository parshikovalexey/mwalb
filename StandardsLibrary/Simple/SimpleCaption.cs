using System.ComponentModel;

namespace StandardsLibrary.Simple
{
    public class SimpleCaption
    {
        [DisplayName("Межстрочный интервал")] public float LineSpacing { get; set; }
        [DisplayName("Отступ слева")] public float LeftIndentation { get; set; }
        [DisplayName("Отступ слева")] public float RightIndentation { get; set; }
        [DisplayName("Нумерация Изображений")] public SimpleCaptionPictureTable Picture { get; set; }
        [DisplayName("Нумерация таблиц")] public SimpleCaptionPictureTable Table { get; set; }
    }
}
