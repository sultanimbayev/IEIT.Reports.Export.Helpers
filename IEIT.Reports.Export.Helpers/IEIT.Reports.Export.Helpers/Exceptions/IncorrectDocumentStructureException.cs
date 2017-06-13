using System;

namespace IEIT.Reports.Export.Helpers.Exceptions
{
    public class IncorrectDocumentStructureException : Exception
    {
        public const string DEFAULT_MESSAGE = @"Нарушена структура документа! Перепроверьте шаблон!";

        public IncorrectDocumentStructureException() : base(DEFAULT_MESSAGE) {}

    }
}