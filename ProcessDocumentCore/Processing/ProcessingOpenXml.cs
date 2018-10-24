using CommonLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        public ResultExecute Processing(Standards designStandard, string filePath)
        {

            return new ResultExecute(){Callbacks = filePath};
        }
    }
}
