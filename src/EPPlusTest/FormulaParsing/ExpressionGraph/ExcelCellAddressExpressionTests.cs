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
        public void VerifyCellAddressExpression_PartlyFixed()
        {
            //Setup
            var f = @"SUM(A$1:$C5:$E$2)";
            var tokens = _tokenizer.Tokenize(f);
            var exps = _graphBuilder.Build(tokens);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            Assert.AreEqual(1, exps.Expressions.Count);

            Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[6].TokenType);
            var result1 = ((CellAddressExpression)exps.Expressions[0].Children[0].Children[0]).Compile();
            var result2 = ((CellAddressExpression)exps.Expressions[0].Children[0].Children[1]).Compile();
            var result3 = ((CellAddressExpression)exps.Expressions[0].Children[0].Children[2]).Compile();
            var range1 = (IRangeInfo)result1.Result;
            var range2 = (IRangeInfo)result2.Result;
            var range3 = (IRangeInfo)result3.Result;

            Assert.AreEqual(range1.Address.FromRow, 1);
            Assert.AreEqual(range1.Address.FromCol, 1);
            Assert.AreEqual(range1.Address.FixedFlag, FixedFlag.FromRowFixed | FixedFlag.ToRowFixed);
            Assert.AreEqual(range2.Address.ToRow, 5);
            Assert.AreEqual(range2.Address.ToCol, 3);
            Assert.AreEqual(range2.Address.FixedFlag, FixedFlag.FromColFixed | FixedFlag.ToColFixed);
            Assert.AreEqual(range3.Address.ToRow, 2);
            Assert.AreEqual(range3.Address.ToCol, 5);
            Assert.AreEqual(range3.Address.FixedFlag, FixedFlag.All);
        }
    }
}
