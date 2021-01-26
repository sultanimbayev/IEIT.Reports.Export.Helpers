using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Styling
{
    public class xlCellStyle
    {
        private Border _border;
        private Fill _fill;
        private NumberingFormat _numFormat;
        private Font _font;
        private CellFormat _cellFormat;
        private xlCellStyleCache _cache;

        public xlCellStyle()
        {
            ResetCache();
        }
        public xlCellStyle(CellFormat cellFormat)
        {
            SetStyleFrom(cellFormat);
        }

        public CellFormat CellFormat
        {
            get => _cellFormat ?? (_cellFormat = new CellFormat());
        }

        public void SetStyleFrom(CellFormat cellFormat)
        {
            var stylesheet = cellFormat.ParentOfType<Stylesheet>();
            if(stylesheet == null)
            {
                throw new InvalidDocumentStructureException("Given CellFormat must be a part of a stylesheet");
            }
            _cellFormat = cellFormat.CloneNode(true) as CellFormat;
            if (cellFormat.BorderId.HasValue)
            {
                _border = stylesheet.Border(cellFormat.BorderId);
            }
            if (cellFormat.FillId.HasValue)
            {
                _fill = stylesheet.Fill(cellFormat.FillId);
            }
            if (cellFormat.NumberFormatId.HasValue)
            {
                _numFormat = stylesheet.NumFormat(cellFormat.NumberFormatId);
            }
            if (cellFormat.FontId.HasValue)
            {
                _font = stylesheet.Font(cellFormat.FontId);
            }
            ResetCache();
            var wbPart = stylesheet.GetWorkbookPart();
            if (stylesheet.GetWorkbookPart() != null)
            {
                _cache.SetStyleIndex(wbPart.Workbook.GetWorkbookId(), (uint)cellFormat.Index());
            }
        }

        public void SetBorder(Border border)
        {
            _border = border;
            if(_cellFormat == null) { _cellFormat = new CellFormat(); }
            CellFormat.ApplyBorder = true;
            ResetCache();
        }

        public void SetFont(Font font)
        {
            _font = font;
            if (_cellFormat == null) { _cellFormat = new CellFormat(); }
            _cellFormat.ApplyFont = true;
            ResetCache();
        }
        public void SetFill(Fill fill)
        {
            _fill = fill;
            if (_cellFormat == null) { _cellFormat = new CellFormat(); }
            _cellFormat.ApplyFill = true;
            ResetCache();
        }
        public void SetNumFormat(NumberingFormat numFormat)
        {
            _numFormat = numFormat;
            if (_cellFormat == null) { _cellFormat = new CellFormat(); }
            _cellFormat.ApplyNumberFormat = true;
            ResetCache();
        }

        public uint GetStyleIndexFor(WorkbookPart wbPart)
        {
            if(_cellFormat == null) { _cellFormat = new CellFormat(); }
            var wbId = wbPart.Workbook.GetWorkbookId();
            if (_cache.ContainsStyleFor(wbId)) { return _cache.GetStyleIndexFor(wbId).Value;}
            var stylesheet = wbPart.GetStylesheet();
            var cellFormat = _cellFormat.CloneNode(true) as CellFormat;
            if(_border != null) { cellFormat.BorderId = stylesheet.Border(_border); }
            if(_fill != null) { cellFormat.FillId = stylesheet.Fill(_fill); }
            if(_font != null) { cellFormat.FontId = stylesheet.Font(_font); }
            if(_numFormat != null) { cellFormat.NumberFormatId = stylesheet.NumFormat(_numFormat); }
            cellFormat.NumberFormatId = _numFormat == null ? null : new UInt32Value(stylesheet.NumFormat(_numFormat));
            var styleIndex = stylesheet.CellFormat(cellFormat);
            _cache.SetStyleIndex(wbId, styleIndex);
            return styleIndex;
        }

        public void ResetCache()
        {
            _cache = new xlCellStyleCache();
        }

    }
}
