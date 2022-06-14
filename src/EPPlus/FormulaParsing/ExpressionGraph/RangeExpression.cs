using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class RangeExpression : GroupExpression
    {
        public RangeExpression(ParsingContext ctx) : base(false,ctx)
        {

        }
        
        public override CompileResult Compile()
        {
            CompileResult cr=null;
            for(int i=0;i<Children.Count-1;i++)
            {
                if (Children[i].Operator == null) return CompileResult.Empty;
                cr = Children[i].Operator.Apply(cr??Children[i].Compile(), Children[i + 1].Compile(), Context);
            }
            return cr;
        }
    }
}
