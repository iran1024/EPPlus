using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class RangeDictionaryTests : TestBase
    {
        //public const int _cellBits = 6;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [TestMethod]
        public void VerifyAddress()
        {
            var rd = new RangeDictionary();

            rd.Merge(new FormulaRangeAddress() { FromRow = 1, ToRow = 5, FromCol = 1, ToCol = 5 });
            rd.Merge(new FormulaRangeAddress() { FromRow = 6, ToRow = 10, FromCol = 1, ToCol = 3 });
            rd.Merge(new FormulaRangeAddress() { FromRow = 1, ToRow = 2, FromCol = 1, ToCol = 3 });
            rd.Merge(new FormulaRangeAddress() { FromRow = 7, ToRow = 8, FromCol = 1, ToCol = 3 });

            Assert.AreEqual(9, rd._addresses[1][0]);
            Assert.AreEqual(9, rd._addresses[2][0]);
            Assert.AreEqual(9, rd._addresses[3][0]);
            Assert.AreEqual(4, rd._addresses[4][0]);
            Assert.AreEqual(4, rd._addresses[5][0]);

        }
    }
}
