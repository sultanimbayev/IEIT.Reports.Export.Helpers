using System;

namespace IEIT.Reports.Export.Helpers.Exceptions
{
    public class InvalidDocumentStructureException : Exception
    {

        public const string DEFAULT_MESSAGE = @"Нарушена структура документа! Перепроверьте шаблон!";

        public InvalidDocumentStructureException(string msg) : base(msg) {}
        public InvalidDocumentStructureException() : base(DEFAULT_MESSAGE) {}

    }
}