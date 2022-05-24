using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class WorksheetNameExpression : Expression
    {
        public WorksheetNameExpression(string expression, ParsingContext ctx) : base(expression, ctx)
        {
        }

        public override bool IsGroupedExpression => false;

        public override CompileResult Compile()
        {
            var sheetName = ExpressionString;
            if(!string.IsNullOrEmpty(sheetName))
            {
                sheetName = sheetName.TrimStart('\'').TrimEnd('\'').Replace("''", "'");
            }
            return new CompileResult(sheetName, DataType.WorksheetName);
        }
    }
}
