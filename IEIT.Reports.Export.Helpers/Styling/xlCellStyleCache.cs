using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IEIT.Reports.Export.Helpers.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Styling
{
    public class xlCellStyleCache
    {
        private Dictionary<string, uint> _stylesIndiciesCache;

        public xlCellStyleCache()
        {
            _stylesIndiciesCache = new Dictionary<string, uint>();
        }

        public bool ContainsStyleFor(string workbookId)
        {
            return _stylesIndiciesCache.ContainsKey(workbookId);
        }

        public uint? GetStyleIndexFor(string workbookId)
        {
            if(!_stylesIndiciesCache.ContainsKey(workbookId))
            {
                return null;
            }
            return _stylesIndiciesCache[workbookId];
        }
        public void SetStyleIndex(string workbookId, uint styleIndex)
        {
            if (!_stylesIndiciesCache.ContainsKey(workbookId))
            {
                _stylesIndiciesCache.Add(workbookId, styleIndex);
                return;
            }
            _stylesIndiciesCache[workbookId] = styleIndex;
        }
    }
}
