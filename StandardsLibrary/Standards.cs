namespace StandardsLibrary
{
    public abstract class Standards
    {
        public abstract string GetFont();
        public abstract int GetFontSize ();
        public abstract float GetLineSpacing ();
        public abstract float GetBeforeSpacing();
        public abstract float GetAfterSpacing();
        public abstract float GetFirstLineIndentation();
        public abstract float GetLeftIndentation();
        public abstract float GetRightIndentation();
        public abstract string GetAlignment ();
        public abstract int GetMarginLeft ();
        public abstract int GetMarginRight ();
        public abstract int GetMarginTop ();
        public abstract int GetMarginBottom ();
        public abstract bool isBold();
        public abstract string GetHeaderColor();
    }
}
