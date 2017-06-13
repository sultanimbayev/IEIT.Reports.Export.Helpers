using System;

namespace IEIT.Reports.Export.Helpers.Exceptions
{
    public class IncompleteActionException : Exception
    {
        public const string DEFAULT_MESSAGE = "Не удалось выполнить операцию.";
        public IncompleteActionException() : base(DEFAULT_MESSAGE){}
        public IncompleteActionException(string actionName) : base(DEFAULT_MESSAGE + $"{ actionName}"){ }
    }
}