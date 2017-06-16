using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers
{
    internal static class Common
    {
        /// <summary>
        /// Регулярное выражения соответствующее адресу ячейки
        /// </summary>
        internal const string RGX_PAT_CA = @"^[a-zA-Z]+\d+$";

        /// <summary>
        /// Регулярное выражение соответствующее ряду адресов ячеек
        /// </summary>
        internal const string RGX_PAT_CA_RANGE = @"^[a-zA-Z]+\d+:[a-zA-Z]+\d+$";
    }
}
