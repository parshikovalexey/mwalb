using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandardsLibrary.Simple
{
  public  class SimpleHeaderGost
    {
        public string Header { get; set; } = String.Empty;
        public Guid GuidGost { get; set; } = Guid.NewGuid();
    }
}
