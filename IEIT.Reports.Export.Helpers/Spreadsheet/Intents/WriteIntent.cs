using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Intents
{
    /// <summary>
    /// Класс для хранения значении для единичного отрабатывания
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Firable<T>
    {
        /// <summary>
        /// Статус отработки значения
        /// </summary>
        public bool Fired;

        /// <summary>
        /// Отрабатываемое значение
        /// </summary>
        private T _value;

        /// <summary>
        /// Отрабатываемое значение. 
        /// При присвоении, значение <see cref="Fired"/> становится false
        /// </summary>
        public T Value
        {
            get { return _value; }
            set { Fired = false; _value = value; }
        }

        /// <summary>
        /// Функция (делегат) отработки значения
        /// </summary>
        private Func<Worksheet, string, T, bool> _fireFunc;
        //private Func<Worksheet, string, string, RunProperties, bool> _appendFireFunc;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="fireFunction">Функция отработки</param>
        public Firable(Func<Worksheet, string, T, bool> fireFunction)
        {
            Fired = true;
            _fireFunc = fireFunction;
        }


        /// <summary>
        /// Вызов отрабатываемой функции. При успешной отработки делегата
        /// свойству <see cref="Fired"/> присваивается значение true
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="cellAddress"></param>
        /// <param name="val"></param>
        /// <returns>true если делегат отработал успешно, false в обратном случае</returns>
        public bool Fire(Worksheet ws, string cellAddress, T val)
        {
            var result = _fireFunc(ws, cellAddress, val);
            Fired = result;
            return result;
        }

    }
    
    /// <summary>
    /// Класс описывающии "Намерение" для записи данных в ячейку или изменение его стиля.
    /// </summary>
    public class WriteIntent
    {

        private Worksheet Worksheet { get; set; }

        private Firable<string> IntendedText { get ; set; }

        private string CellAddress { get; set; }

        private Firable<UInt32Value> IntendedStyle { get; set; }
        
        /// <summary>
        /// Создает "намерение" для изменения своиств ячейки.
        /// </summary>
        /// <param name="ws">Рабочий лист в котором будут изменения</param>
        public WriteIntent(Worksheet ws)
        {
            IntendedText = new Firable<string>(WriterActions._writeAny);
            IntendedStyle = new Firable<UInt32Value>(WriterActions._setStyle);
            Worksheet = ws;
        }

        /// <summary>
        /// Создает "намерение" для изменения своиств ячейки 
        /// с переопределением поведения записи значения в ячейку.
        /// Используйте этот конструктор только в случае если конструктор 
        /// <see cref="WriteIntent(Worksheet)"/> не дает нужных результатов
        /// </summary>
        /// <param name="ws">Рабочий лист в котором будут изменения</param>
        /// <param name="writeDeleg">
        /// Делегат для записи значении в ячейку. 
        /// По сути это определяет то, как будет записыватся значение в ячейку.
        /// </param>
        public WriteIntent(Worksheet ws, Func<Worksheet, string, string, bool> writeDeleg) : this(ws)
        {
            IntendedText = new Firable<string>(writeDeleg);
        }

        /// <summary>
        /// Запись в ячейку или присвоения стиля ячейки с указанным адресом
        /// </summary>
        /// <param name="cellAddress">Адрес ячейки свойства которой требуется изменить</param>
        /// <returns>WriteIntent для изменения своиств ячейки</returns>
        public WriteIntent To(string cellAddress)
        {
            CellAddress = cellAddress;
            if (canFire()) { fireAll(); };
            return this;
        }

        /// <summary>
        /// Запись в ячейку или присвоения стиля ячейки с указанным адресом
        /// </summary>
        /// <param name="columnNum">Номер колонки ячейки</param>
        /// <param name="rowNum">Номер строки ячейки</param>
        /// <returns>WriteIntent для изменения своиств ячейки</returns>
        public WriteIntent To(int columnNum, int rowNum)
        {
            CellAddress = Utils.ToColumnName((uint)columnNum) + rowNum.ToString();
            if (canFire()) { fireAll(); };
            return this;
        }

        /// <summary>
        /// Запись в ячейку или присвоения стиля ячейки с указанным адресом
        /// </summary>
        /// <param name="columnNum"></param>
        /// <param name="rowNum"></param>
        /// <returns>WriteIntent для изменения своиств ячейки</returns>
        public WriteIntent To(uint columnNum, uint rowNum)
        {
            CellAddress = Utils.ToColumnName(columnNum) + rowNum.ToString();
            if (canFire()) { fireAll(); };
            return this;
        }

        /// <summary>
        /// Присваивает стиль всей ячейке
        /// </summary>
        /// <param name="styleIndex"></param>
        /// <returns>возвращает "Намерение" изменения данных ячейки</returns>
        public WriteIntent WithStyle(UInt32Value styleIndex)
        {
            IntendedStyle.Value = styleIndex;
            if (canFire()) { fireAll(); };
            return this;
        }

        /// <summary>
        /// Задать текст указанной ячейке
        /// </summary>
        /// <param name="text">Новый текст ячейки</param>
        /// <returns>возвращает "Намерение" изменения данных ячейки</returns>
        public WriteIntent WithText(string text)
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