using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class OptimizedDependencyChainTests : TestBase
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

            parsingContext.Scopes.NewScope(new FormulaRangeAddress() { WorksheetIx = 1, FromRow = 1, FromCol = 1, ToRow = 1, ToCol = 1 });
            parsingContext.ExcelDataProvider = _excelDataProvider;
            _graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveWorkbook("DependencyChain.xlsx", _package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyCellAddressExpression_NonFixed()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {                
                var ws = p.Workbook.Worksheets[0];
                var dp=OptimizedDependencyChainFactory.Create(p.Workbook, new ExcelCalculationOption(){ });
            }
        }
        [TestMethod]
        public void VerifyCellAddressExpression_Range()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var dp = OptimizedDependencyChainFactory.Create(ws, new ExcelCalculationOption() { });
            }
        }

    }
}
