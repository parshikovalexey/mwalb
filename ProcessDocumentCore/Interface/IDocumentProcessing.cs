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
        ResultExecute Processing(GostModel gostModel, string filePath);
    }
}
