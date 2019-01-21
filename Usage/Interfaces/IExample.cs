using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Usage.Interfaces
{
    public interface IExample
    {
        void Execute(SpreadsheetDocument doc);
    }
}
