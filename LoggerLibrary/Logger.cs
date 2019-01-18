using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace LoggerLibrary
{
  public  static class Logger
    {
        public static NLog.Logger Write()
        {
            return LogManager.GetLogger("Default");
        }
    }
}
