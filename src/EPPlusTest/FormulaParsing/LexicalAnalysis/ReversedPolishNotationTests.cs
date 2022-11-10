using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class ReversedPolishNotationTests
    {
        ExcelPackage _package;
        ParsingContext _parsingContext;
        RpnExpressionGraph _graph;
        private ISourceCodeTokenizer _tokenizer;
        private RpnFunctionCompilerFactory _functionCompilerFactory;
       [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _parsingContext = ParsingContext.Create(_package);

            var dataProvider = new EpplusExcelDataProvider(_package, _parsingContext);
            _parsingContext.ExcelDataProvider = dataProvider;
            _parsingContext.NameValueProvider = new EpplusNameValueProvider(dataProvider);
            _parsingContext.RangeAddressFactory = new RangeAddressFactory(dataProvider, _parsingContext);

            _graph = new RpnExpressionGraph(_parsingContext);
            _tokenizer = OptimizedSourceCodeTokenizer.Default;

            var ws1 = _package.Workbook.Worksheets.Add("Sheet1");
            ws1.Cells["A1"].Value = 1;
            ws1.Cells["B1"].Value = 2;
            ws1.Cells["C1"].Value = 3;

            ws1.Cells["A2"].Value = 10;
            ws1.Cells["B2"].Value = 20;
            ws1.Cells["C2"].Value = 30;

            _package.Workbook.Names.AddValue("WorkbookDefinedNameValue", 1);
            ws1.Names.AddValue("WorksheetDefinedNameValue", "Name Value");

            var ws2 = _package.Workbook.Worksheets.Add("Sheet2");
            ws2.Cells["A1:A3"].Value = 4;
        }

        [TestMethod]
        public void Calculate_NumericExpression1()
        {
            var formula = "3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual(3.001953125D, cr.ResultNumeric);
        }
        [TestMethod]
        public void Calculate_NumericExpression2()
        {
            var formula = "(( 1 -(- 2)-( 3 + 4 + 5 ))/( 6 + 7 * 8 - 9) * 10 )";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual(-1.6981132075471697D, cr.ResultNumeric);
        }
        [TestMethod]
        public void Calculate_NumericExpression3()
        {
            var formula = "( 1 + 2 ) * ( 3 / 4 ) ^ ( 5 + 6 )";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual(0.12670540809631348D, cr.ResultNumeric);
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithFunctions()
        {
            var formula = "sin(max((( 2 + 2 ) / 2), (3 * 3) / 3) / 3 * pi())";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual(3.231085104332676E-15, cr.ResultNumeric);
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithAddresses1()
        {
            var formula = "A1 + B1 * C1 / ( 1 - 5 ) ^ 2 ^ 3";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual(1.00146484375000, cr.ResultNumeric);
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithAddresses2()
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create("sheet1", 4, 1);
            using (_parsingContext.Scopes.NewScope(rangeAddress))
            {
                var formula = "(SUM(Sheet1!A1:C1)+1) * 3";
                var tokens = _tokenizer.Tokenize(formula);
                var exps = _graph.CreateExpressionList(tokens);
                var cr = _graph.Execute(exps);
                
                Assert.AreEqual(21, cr.ResultNumeric);
            }
        }
        [TestMethod]
        public void Calculate_NumericExpressionMultiplyTwoRanges()
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create("sheet1", 4, 1);
            using (_parsingContext.Scopes.NewScope(rangeAddress))
            {
                for (int i = 0; i < 1000000; i++)
                {
                    var formula = "SUM(A1:B1+A2:B2)+1";
                    var tokens = _tokenizer.Tokenize(formula);
                    var exps = _graph.CreateExpressionList(tokens);
                    var cr = _graph.Execute(exps);

                    Assert.AreEqual(34, cr.ResultNumeric);
                }
            }
        }
        [TestMethod]
        public void Calculate_Concat_Strings()
        {
            var formula = "\"Test\" & \" \" & \"2\"";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual("Test 2", cr.Result);
        }
        [TestMethod]
        public void Calculate_Array()
        {
            var formula = "Sum({1,2;3,4})";
            var tokens = _tokenizer.Tokenize(formula);
            var exps = _graph.CreateExpressionList(tokens);
            var cr = _graph.Execute(exps);
            Assert.AreEqual("Test 2", cr.Result);
        }

        [TestMethod]
        public void Calculate_Worksheet_NameFixedValue()
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create("sheet1", 4, 1);
            using (_parsingContext.Scopes.NewScope(rangeAddress))
            {
                var formula = "Sheet1!WorksheetDefinedNameValue";
                var tokens = _tokenizer.Tokenize(formula);
                var exps = _graph.CreateExpressionList(tokens);
                var cr = _graph.Execute(exps);
                Assert.AreEqual("Name Value", cr.Result);
            }
        }
        [TestMethod]
        public void Calculate_Workbook_NameFixedValue()
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create("sheet1", 4, 1);
            using (_parsingContext.Scopes.NewScope(rangeAddress))
            {
                var formula = "WorkbookDefinedNameValue";
                var tokens = _tokenizer.Tokenize(formula);
                var exps = _graph.CreateExpressionList(tokens);
                var cr = _graph.Execute(exps);
                Assert.AreEqual(1, cr.Result);
            }
        }
        [TestMethod]
        public void Calculate_NonExisting_Worksheet_NameFixedValue()
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create("sheet1", 4, 1);
            using (_parsingContext.Scopes.NewScope(rangeAddress))
            {
                var formula = "NonExistingSheet!WorksheetDefinedNameValue";
                var tokens = _tokenizer.Tokenize(formula);
                var exps = _graph.CreateExpressionList(tokens);
                var cr = _graph.Execute(exps);
                Assert.IsInstanceOfType(cr.Result, typeof(ExcelErrorValue));
                Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)cr.Result).Type);
            }
        }
    }
}
