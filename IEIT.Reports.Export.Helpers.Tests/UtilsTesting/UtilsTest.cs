using System;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;

namespace IEIT.Reports.Export.Helpers.Tests
{
    [TestFixture]
    class UtilsTest
    {
        [TestCase(1, "A1")]
        [TestCase(22, "A22")]
        [TestCase(1, "AA1")]
        [TestCase(22, "AA22")]
        public void ToRowNumTest(int expected, string input)
        {
            Assert.AreEqual(expected, Utils.ToRowNum(input));
        }

        [TestCase(1, "A1")]
        [TestCase(26, "Z22")]
        [TestCase(27, "AA1")]
        [TestCase(27, "AA22")]
        [TestCase(27, "22")]
        public void ToColumnNum(int expected, string address)
        {
            Assert.AreEqual(expected, Utils.ToColumnNum(address));
        }

        [TestCase("A", "A1")]
        [TestCase("B", "B1")]
        [TestCase("ABC", "ABC112")]
        [TestCase("ABC", "ABC112")]
        [TestCase("A", "1")]
        [TestCase("B", "2")]
        [TestCase("C", "3")]
        [TestCase("D", "4")]
        [TestCase("Z", "26")]
        [TestCase("AA", "27")]
        [TestCase("AB", "28")]
        [TestCase("AC", "29")]
        [TestCase("AD", "30")]
        public void ToColumnName(string expected, string address)
        {
            Assert.AreEqual(expected, Utils.ToColumnName(address));
        }

        [TestCase("A",1)]
        [TestCase("B",2)]
        [TestCase("C",3)]
        [TestCase("D",4)]
        [TestCase("Z", 26)]
        [TestCase("AA", 27)]
        [TestCase("AB", 28)]
        [TestCase("AC", 29)]
        [TestCase("AD", 30)]
        public void ToColumnNameWithNumber(string expected, int input)
        {
            Assert.AreEqual(expected, Utils.ToColumnName(input));
        }

        [TestCase(true, "A1")]
        [TestCase(false, "")]
        [TestCase(false, null)]
        [TestCase(false, "   ")]
        [TestCase(false, "РУССКИЕЗАГЛАВНЫЕ123")]
        [TestCase(false, "ASDF")]
        [TestCase(false, "123")]
        [TestCase(false, "ASDF123ASDF")]
        [TestCase(true, "asdf123")]
        [TestCase(true, "   ASDF123   ")]
        [TestCase(false, "!№   ASDF123   ")]
        public void IsCellAddress(bool expected, string input)
        {
            Assert.AreEqual(expected, Utils.IsCellAddress(input));
        }
    }
}
