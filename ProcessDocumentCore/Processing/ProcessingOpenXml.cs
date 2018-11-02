using CommonLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
using DocumentFormat.OpenXml.Packaging;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        public ResultExecute Processing(Standards designStandard, string filePath)
        {
            try
            {
                WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true);
                return new ResultExecute() { Callbacks = wordDocument, ErrorMsg = "", StatusExecute = ResultExecute.Status.Success };
            }
            catch (Exception ex)
            {
                return new ResultExecute() { Callbacks = null, ErrorMsg = ex.ToString(), StatusExecute = ResultExecute.Status.Error };
            }
        }
    }
}
