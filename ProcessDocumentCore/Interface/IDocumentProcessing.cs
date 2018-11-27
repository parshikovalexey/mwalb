using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommonLibrary;
using StandardsLibrary;

namespace ProcessDocumentCore.Interface
{
    public interface IDocumentProcessing
    {
        ResultExecute Processing(Standards designStandard, string filePath, GostModel gostModel);
        ResultExecute Processing(GostModel designStandard, string filePath);
    }
}
