using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.OperatorsTests
{

    [TestClass]
    public class RangeOperationsOperatorTests
    {
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _context = ParsingContext.Create(_package);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        private ParsingContext _context;
        private ExcelPackage _package;

        [TestMethod]
        public void ShouldSetNAerrorWithMultipleCellRangesOfDifferentSize()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);
            r2.SetValue(1, 0, 2);
            r2.SetValue(1, 1, 3);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(5d, range.GetValue(1, 0));
            Assert.AreEqual(5d, range.GetValue(1, 1));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(2, 0));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(2, 1));
        }

        [TestMethod]
        public void ShouldCalculateWithSameColumnSize()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 1);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(1, 0, 2);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(3d, range.GetValue(1, 0));
            Assert.AreEqual(4d, range.GetValue(1, 1));
            Assert.AreEqual(4d, range.GetValue(2, 0));
            Assert.AreEqual(5d, range.GetValue(2, 1));
        }

        [TestMethod]
        public void ShouldCalculateWithSameRowSize()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(2d, range.GetValue(0, 1));
            Assert.AreEqual(5d, range.GetValue(1, 0));
            Assert.AreEqual(5d, range.GetValue(1, 1));
        }
    }
}
