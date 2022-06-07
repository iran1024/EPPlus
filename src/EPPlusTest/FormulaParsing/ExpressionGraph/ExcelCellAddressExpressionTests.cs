using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExcelCellAddressTests : TestBase
    {
        static ExcelPackage _package;
        static EpplusExcelDataProvider _excelDataProvider;
        static ExpressionGraphBuilder _graphBuilder;
        static ExcelWorksheet _ws;
        internal static ISourceCodeTokenizer _tokenizer = OptimizedSourceCodeTokenizer.Default;
        static ExpressionCompiler _compiler;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _package = new ExcelPackage();

            _ws = _package.Workbook.Worksheets.Add("Sheet1");
            LoadTestdata(_ws);
            var tbl = _ws.Tables.Add(_ws.Cells["A1:E101"], "MyTable");
            tbl.ShowTotal = true;
            _excelDataProvider = new EpplusExcelDataProvider(_package);
            var parsingContext = ParsingContext.Create(_package);
            _compiler = new ExpressionCompiler(parsingContext);

            parsingContext.Scopes.NewScope(new FormulaRangeAddress() { WorksheetIx=1, FromRow = 1, FromCol = 1, ToRow = 1,ToCol = 1 });
            parsingContext.ExcelDataProvider = _excelDataProvider;
            _graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveWorkbook("CellAddressExpression.xlsx", _package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyCellAddressExpression_NonFixed()
        {
            //Setup
            var f = @"SUM(A1:C5)";
            var tokens = _tokenizer.Tokenize(f);
            var exps = _graphBuilder.Build(tokens);

            //Assert
            Assert.AreEqual(6, tokens.Count);
            Assert.AreEqual(1, exps.Expressions.Count);

            Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            var result1 = ((CellAddressExpression)exps.Expressions[0].Children[0].Children[0]).Compile();
            var result2 = ((CellAddressExpression)exps.Expressions[0].Children[0].Children[1]).Compile();
            var range1 = (IRangeInfo)result1.Result;
            var range2 = (IRangeInfo)result2.Result;

            Assert.AreEqual(range1.Address.FromRow, 1);
            Assert.AreEqual(range1.Address.FromCol, 1);
            Assert.AreEqual(range1.Address.FixedFlag, FixedFlag.None);
            Assert.AreEqual(range2.Address.ToRow, 5);
            Assert.AreEqual(range2.Address.ToCol, 3);
            Assert.AreEqual(range2.Address.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_MultiColon()
        {
            //Setup
            var f = @"Sheet1!A1:C5:E2";
            var tokens = _tokenizer.Tokenize(f);
            var exps = _graphBuilder.Build(tokens);
            var result = _compiler.Compile(exps.Expressions);
            _package.Workbook.Worksheets[0].Cells["H5"].Formula = f;

            //Assert
            Assert.AreEqual(7, tokens.Count);
            Assert.AreEqual(3, exps.Expressions.Count);

            Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[6].TokenType);
            var range = (IRangeInfo)result.Result;

            Assert.AreEqual(range.Address.ExternalReferenceIx, -1);
            Assert.AreEqual(range.Address.WorksheetIx, 0);

            Assert.AreEqual(range.Address.FromRow, 1);
            Assert.AreEqual(range.Address.FromCol, 1);
            Assert.AreEqual(range.Address.ToRow, 5);
            Assert.AreEqual(range.Address.ToCol, 5);
            Assert.AreEqual(range.Address.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_Fixed()
        {
            //Setup
            var f = @"[0]Sheet1!A1:C5:E2";
            var tokens = _tokenizer.Tokenize(f);
            var exps = _graphBuilder.Build(tokens);
            var result = _compiler.Compile(exps.Expressions);
            _package.Workbook.Worksheets[0].Cells["H5"].Formula = f;

            //Assert
            Assert.AreEqual(10, tokens.Count);
            Assert.AreEqual(3, exps.Expressions.Count);

            Assert.AreEqual(TokenType.CellAddress, tokens[5].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[7].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[9].TokenType);
            var range = (IRangeInfo)result.Result;

            Assert.AreEqual(range.Address.ExternalReferenceIx, -1);
            Assert.AreEqual(range.Address.WorksheetIx, 0);

            Assert.AreEqual(range.Address.FromRow, 1);
            Assert.AreEqual(range.Address.FromCol, 1);
            Assert.AreEqual(range.Address.ToRow, 5);
            Assert.AreEqual(range.Address.ToCol, 5);
            Assert.AreEqual(range.Address.FixedFlag, FixedFlag.None);
        }
    }
}
