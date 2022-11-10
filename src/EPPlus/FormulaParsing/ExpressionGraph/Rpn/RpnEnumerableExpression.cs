/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    /// <summary>
    /// This expression represents a literal array where rows and cols are separated with comma and semicolon.
    /// </summary>
    public class RpnEnumerableExpression : RpnExpression
    {
        private readonly List<List<object>> _matrix;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="expressionCompiler"></param>
        /// <param name="ctx"></param>
        public RpnEnumerableExpression(List<List<object>> matrix, ParsingContext ctx)
            : base(ctx)
        {
            _matrix = matrix;
        }
        internal override ExpressionType ExpressionType => ExpressionType.Enumerable;

        ///// <summary>
        ///// Compiles the expression into a <see cref="CompileResult"/>
        ///// </summary>
        ///// <returns></returns>
        //public override CompileResult Compile()
        //{
        //    var rangeDef = GetRangeDefinition();
        //    var result = new InMemoryRange(rangeDef);
        //    var rowIx = 0;
        //    var colIx = 0;
        //    for(var ix = 0; ix < _matrix.Count; ix++)
        //    {
        //        var childExpression = Children[ix];
        //        var childResult = _expressionCompiler.Compile(new List<Expression> { childExpression }).Result;
        //        result.SetValue(rowIx, colIx, childResult);
        //        if (ix < _separators.Count)
        //        {
        //            if (_separators[ix].TokenTypeIsSet(TokenType.SemiColon))
        //            {
        //                rowIx++;
        //                colIx = 0;
        //            }
        //            else if (_separators[ix].TokenTypeIsSet(TokenType.Comma))
        //            {
        //                colIx++;
        //            }
        //        }
        //    }
        //    return new CompileResult(result, DataType.ExcelRange);
        //}

        //private RangeDefinition GetRangeDefinition()
        //{
        //    short nCols = 1;
        //    var ix = 0;
        //    while(ix < _separators.Count && _separators[ix].TokenTypeIsSet(TokenType.Comma))
        //    {
        //        ix++;
        //        nCols++;
        //    }
        //    var nRows = 1;
        //    nRows += _separators.Count(x => x.TokenTypeIsSet(TokenType.SemiColon));
        //    return new RangeDefinition(nRows, nCols);
        //}
    }
}
