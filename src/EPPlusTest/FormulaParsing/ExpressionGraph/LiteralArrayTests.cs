using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class LiteralArrayTests
    {
        private ExcelDataProvider _provider;
        private ExcelPackage _package;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _provider = new EpplusExcelDataProvider(_package);
        }

        [TestMethod]
        public void ShouldHandleLiteralArrayMulti()
        {
            var tokens = SourceCodeTokenizer.Optimized.Tokenize("{1,2,3;4,5,6}");
            Assert.AreEqual(13, tokens.Count());
            var graphBuilder = new ExpressionGraphBuilder(_provider, ParsingContext.Create(_package));
            var exp = graphBuilder.Build(tokens);
            var result = exp.Current.Compile();
            Assert.IsNotNull(result);
            Assert.AreEqual(DataType.ExcelRange, result.DataType);
            Assert.IsInstanceOfType(result.Result, typeof(InMemoryRange));
            var range = (InMemoryRange)result.Result;
            Assert.AreEqual(6, range.GetNCells());
            Assert.AreEqual(3, range.Size.NumberOfCols);
            Assert.AreEqual(2, range.Size.NumberOfRows);
        }

        [TestMethod]
        public void ShouldHandleLiteralArraySingleRow()
        {
            var tokens = SourceCodeTokenizer.Optimized.Tokenize("{1,2,3,4,5,6}");
            Assert.AreEqual(13, tokens.Count());
            var graphBuilder = new ExpressionGraphBuilder(_provider, ParsingContext.Create(_package));
            var exp = graphBuilder.Build(tokens);
            var result = exp.Current.Compile();
            Assert.IsNotNull(result);
            Assert.AreEqual(DataType.ExcelRange, result.DataType);
            Assert.IsInstanceOfType(result.Result, typeof(InMemoryRange));
            var range = (InMemoryRange)result.Result;
            Assert.AreEqual(6, range.GetNCells());
            Assert.AreEqual(6, range.Size.NumberOfCols);
            Assert.AreEqual(1, range.Size.NumberOfRows);
        }

        [TestMethod]
        public void ShouldHandleLiteralArraySingleCol()
        {
            var tokens = SourceCodeTokenizer.Optimized.Tokenize("{1;2;3;4;5;6}");
            Assert.AreEqual(13, tokens.Count());
            var graphBuilder = new ExpressionGraphBuilder(_provider, ParsingContext.Create(_package));
            var exp = graphBuilder.Build(tokens);
            var result = exp.Current.Compile();
            Assert.IsNotNull(result);
            Assert.AreEqual(DataType.ExcelRange, result.DataType);
            Assert.IsInstanceOfType(result.Result, typeof(InMemoryRange));
            var range = (InMemoryRange)result.Result;
            Assert.AreEqual(6, range.GetNCells());
            Assert.AreEqual(1, range.Size.NumberOfCols);
            Assert.AreEqual(6, range.Size.NumberOfRows);
        }
    }
}
