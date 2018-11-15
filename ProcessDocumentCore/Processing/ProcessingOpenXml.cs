using CommonLibrary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
using System.IO;
using System.Linq;
using ProcessDocumentCore.Models;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        private Standards _designStandard;
        private const string ExtensionDoc = ".docx";
        public ResultExecute Processing(Standards designStandard, string filePath)
        {
            _designStandard = designStandard;
            PathHelper.ClearTmpDirectory();
            Stream stream = null;
            WordprocessingDocument wordDoc = null;
            var pathToSaveObj = GetPathToSaveObj();
            try
            {
                File.Copy(filePath, pathToSaveObj);
                stream = File.Open(pathToSaveObj, FileMode.Open);

                using (wordDoc = WordprocessingDocument.Open(stream, true)){
                    string docText;
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    var body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        bool isNeedChangeStyleForParagraph = false;
                        if (para.Any(p => p is BookmarkStart) && para.Any(p => p is BookmarkEnd))
                        {
                            foreach (var openXmlElement1 in para.Where(x => x is Run).ToList())
                            {
                                var openXmlElement = (Run)openXmlElement1;
                                //Определяем есть ли нумерованный список для задания стиля всему параграфу, т.к. на него распотсраняется отдельный стиль
                                var hasNumberingProperties = para.ParagraphProperties.FirstOrDefault(o =>
                                    o.GetType() == typeof(NumberingProperties));
                                if (hasNumberingProperties != null) isNeedChangeStyleForParagraph = true;

                                RunProperties runProperties = openXmlElement.RunProperties;
                                if (runProperties == null) continue;
                                var paragraphStyles = new ParagraphStyles(_designStandard);
                                paragraphStyles.SetParagraphStylesForRunProperties(runProperties);
                            }
                        }
                        if (isNeedChangeStyleForParagraph) SetParagraphStyle(para);

                    }

                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }

                    wordDoc.Close();
                    stream.Close();
                }
                return new ResultExecute() { Callbacks = pathToSaveObj };
            }
            catch (Exception ex)
            {
                wordDoc?.Close();
                stream?.Close();
                return new ResultExecute().OnError(ex.Message);
            }
        }

        private void SetParagraphStyle(Paragraph para) {
            try {
                //задаем стили для параграфа, т.к. если в параграфе имеется нумерация, то стиль берется общий для параграфа
                if (para.ParagraphProperties?.ParagraphMarkRunProperties != null) {

                    //Подгрузили стили
                    var paragraphStyles = new ParagraphStyles(_designStandard);
                    //Получили коллекцию элементов
                    var openXmlElements = paragraphStyles.GetCollectionOfElements();
                    //Убили лишнее из списка
                    paragraphStyles.RemoveChildElementFromParagraph(para);
                    //Добавили нашу коллекцию в стили параграфа
                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(openXmlElements);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }
}
