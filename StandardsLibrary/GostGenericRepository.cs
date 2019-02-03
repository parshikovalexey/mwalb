﻿using DocumentFormat.OpenXml.Wordprocessing;
using HelperLibrary;
using StandardsLibrary.Simple;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OpenXmlHelperLibrary;

namespace StandardsLibrary
{
    public partial class GostGenericRepository<T> where T : GostModel
    {
        private readonly GostModel _model = null;

        private readonly Dictionary<int, SimpleNumberingLevel> _numberingDictionary = new Dictionary<int, SimpleNumberingLevel>();

        public GostGenericRepository(GostModel @base)
        {
            _model = @base;

            if (_model.Numbering != null)
            {
                _numberingDictionary.Clear();
                _numberingDictionary.Add(0, _model?.Numbering?.Level1 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(1, _model?.Numbering?.Level2 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(2, _model?.Numbering?.Level3 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(3, _model?.Numbering?.Level4 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(4, _model?.Numbering?.Level5 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(5, _model?.Numbering?.Level6 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(6, _model?.Numbering?.Level7 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(7, _model?.Numbering?.Level8 ?? new SimpleNumberingLevel());
                _numberingDictionary.Add(8, _model?.Numbering?.Level9 ?? new SimpleNumberingLevel());
            }

        }

        /// <summary>
        /// Возвращает Размер шрифта
        /// </summary>
        /// <param name="typeStyle"></param>
        /// <returns>Вернет число или null если параметр отсуствует</returns>
        public int? GetFontSize(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.FontSize;

        public string GetFont(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.Font;

        public float? GetLineSpacing(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.LineSpacing;

        public string GetAlignment(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.Alignment;

        public float? GetBeforeSpacing(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.BeforeSpacing;

        public float? GetAfterSpacing(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.AfterSpacing;

        public float? GetFirstLineIndentation(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.FirstLineIndentation;

        public float? GetLeftIndentation(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.LeftIndentation;

        public float? GetRightIndentation(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.RightIndentation;

        public int? GetMarginLeft(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.MarginLeft.ToMargins();

        public int? GetMarginRight(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.MarginRight.ToMargins();

        public int? GetMarginTop(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.MarginTop.ToMargins();

        public int? GetMarginBottom(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.MarginBottom.ToMargins();

        public bool? GetBold(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.Bold;

        public string GetColor(CommonGost.StyleTypeEnum typeStyle) => GetDirectoryStyle(typeStyle)?.Color;

        private SimpleStyle GetDirectoryStyle(CommonGost.StyleTypeEnum typeStyle)
        {
            switch (typeStyle)
            {
                case CommonGost.StyleTypeEnum.GlobalText:
                    return _model?.GlobalText;
                case CommonGost.StyleTypeEnum.Image:
                    return _model?.Image;
                case CommonGost.StyleTypeEnum.Headline:
                    return _model?.Headline;
                default:
                    return null;
            }
        }



        public NumberFormatValues GetNumberingFormat(int level)
        {
            return _numberingDictionary.ContainsKey(level) ? _numberingDictionary[level].NumberingFormat.GetNumberingFormat() : NumberFormatValues.Decimal;
        }
        public string GetNumberingLevelText(int level)
        {
            if (_numberingDictionary.ContainsKey(level))
            {
                return _numberingDictionary[level].LevelText;
            }
            else
            {
                var levelText = string.Empty;
                for (var j = 0; j < level+1; j++)
                {
                    levelText += $"%{j + 1}.";
                }
                return $"{levelText} ";
            }
        }

        public string GetNumberingIndentationLeft(int level)
        {
            float indentationleft = 720;
            float nextIndentationleft = 720;

            if (_model.Numbering.LeftNextIndentation>0) nextIndentationleft = _model.Numbering.LeftIndentation;

            if (_model.Numbering.LeftIndentation > 0) indentationleft = _model.Numbering.LeftIndentation;

            var nextLevel = (level + 1) * nextIndentationleft;
            return level > 0
                ? nextLevel < 1 ? indentationleft.ToString() : nextLevel.ToString()
                : indentationleft.ToString();
        }





    }
}
