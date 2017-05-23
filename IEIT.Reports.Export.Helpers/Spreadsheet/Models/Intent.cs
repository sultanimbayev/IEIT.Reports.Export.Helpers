using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Models
{
    public class Firable<T>
    {
        public bool Fired;

        private T _value;
        public T Value
        {
            get { return _value; }
            set { Fired = false; _value = value; }
        }

        private Func<Worksheet, string, T, bool> _fireFunc;

        public Firable(Func<Worksheet, string, T, bool> fireFunction)
        {
            Fired = true;
            _fireFunc = fireFunction;
        }
        public bool Fire(Worksheet ws, string cellAddress, T val)
        {
            var result = _fireFunc(ws, cellAddress, val);
            Fired = true;
            return result;
        }
    }
    
    public class Intent
    {

        private Worksheet Worksheet { get; set; }

        private Firable<string> IntendedText { get ; set; }

        private string CellAddress { get; set; }

        private Firable<UInt32Value> IntendedStyle { get; set; }
        
        public Intent(Worksheet ws, Func<Worksheet, string, string, bool> writeText, Func<Worksheet, string, UInt32Value, bool> setStyle)
        {
            IntendedText = new Firable<string>(writeText);
            IntendedStyle = new Firable<UInt32Value>(setStyle);
            Worksheet = ws;
        }

        public Intent To(string cellAddress)
        {
            CellAddress = cellAddress;
            if (canFire()) { fireAll(); };
            return this;
        }

        public Intent To(int columnNum, int rowNum)
        {
            CellAddress = Utils.ToColumnName(columnNum) + rowNum.ToString();
            if (canFire()) { fireAll(); };
            return this;
        }

        public Intent WithStyle(UInt32Value styleIndex)
        {
            IntendedStyle.Value = styleIndex;
            if (canFire()) { fireAll(); };
            return this;
        }

        public Intent WithText(string text)
        {
            IntendedText.Value = text;
            if (canFire()) { fireAll(); };
            return this;
        }

        private bool canFire()
        {
            if(CellAddress == null || CellAddress == string.Empty) { return false; }
            return true;
        }

        private void fireAll()
        {
            if (!IntendedText.Fired) { IntendedText.Fire(Worksheet, CellAddress, IntendedText.Value); }
            if (!IntendedStyle.Fired) { IntendedStyle.Fire(Worksheet, CellAddress, IntendedStyle.Value); }
        }

    }
}