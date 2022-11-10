/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2022         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Drawing2D;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal class RpnExpressionGraph
    {
        private ParsingContext _parsingContext;
        private RpnFunctionCompilerFactory _functionCompilerFactory;

        internal RpnExpressionGraph(ParsingContext parsingContext)
        {
            _parsingContext = parsingContext;
            _functionCompilerFactory = new RpnFunctionCompilerFactory(_parsingContext.Configuration.FunctionRepository, _parsingContext);
        }

        //https://www.andreinc.net/2010/10/05/converting-infix-to-rpn-shunting-yard-algorithm
        //https://en.wikipedia.org/wiki/Shunting_yard_algorithm#The_algorithm_in_detail
        public List<Token> CreateExpressionList(IList<Token> tokens)
        {
            var operators = OperatorsDict.Instance;
            Stack<Token> operatorStack = new Stack<Token>();
            var expressions = new List<Token>();
            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                switch (token.TokenType)
                {
                    case TokenType.OpeningParenthesis:
                        operatorStack.Push(token);
                        break;
                    case TokenType.ClosingParenthesis:
                        if (operatorStack.Count > 0)
                        {
                            var o = operatorStack.Pop();
                            while (o.TokenType != TokenType.OpeningParenthesis)
                            {
                                expressions.Add(o);
                                if (operatorStack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = operatorStack.Pop();
                            }
                            if (operatorStack.Count > 0 && operatorStack.Peek().TokenType == TokenType.Function)
                            {
                                expressions.Add(operatorStack.Pop());
                            }
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if (operatorStack.Count > 0)
                        {
                            var o2 = operatorStack.Peek();
                            while (o2.TokenType == TokenType.Operator &&
                                operators[o2.Value].Precedence <= operators[token.Value].Precedence)
                            {
                                expressions.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                o2 = operatorStack.Peek();
                            }
                        }
                        operatorStack.Push(token);
                        break;

                    case TokenType.Function:
                        expressions.Add(new Token("", TokenType.StartFunctionArguments));
                        operatorStack.Push(token);
                        break;
                    case TokenType.Comma:
                    case TokenType.SemiColon:
                        while (operatorStack.Peek().TokenType == TokenType.Operator)
                        {
                            expressions.Add(operatorStack.Pop());
                        }
                        break;
                    default:
                        expressions.Add(token);
                        break;
                }

            }

            while (operatorStack.Count > 0)
            {
                expressions.Add(operatorStack.Pop());
            }

            return expressions;
        }

        public abstract class TokenResult
        {
            public Token Token;
            public abstract void Negate();
            public abstract object Value { get; }
            public abstract void ApplyOperator(Token op, TokenResult tr);
        }
        public class FunctionArgumentTokenResult : TokenResult
        {
            public override void ApplyOperator(Token op, TokenResult tr)
            {
                throw new NotImplementedException();
            }
            public override void Negate()
            {
                throw new NotImplementedException();
            }
            public override object Value => throw new NotImplementedException();
        }
        public class TokenResultDouble : TokenResult
        {
            public TokenResultDouble(Token t, double v)
            {
                Token = t;
                ValueDouble = v;
            }
            public double ValueDouble;
            public override object Value
            {
                get
                {
                    return ValueDouble;
                }
            }
            public override void Negate()
            {
                ValueDouble = -ValueDouble;
            }
            public override void ApplyOperator(Token op, TokenResult tr)
            {
                double v = 0;
                if (tr.Token.TokenType == TokenType.Decimal ||
                   tr.Token.TokenType == TokenType.Integer)
                {
                    v = ((TokenResultDouble)tr).ValueDouble;
                }
                else if (tr.Token.TokenType == TokenType.Boolean)
                {

                }
                else if (tr.Token.TokenType == TokenType.CellAddress ||
                        tr.Token.TokenType == TokenType.ExcelAddress)
                {

                }

                switch (op.Value)
                {
                    case "+":
                        ValueDouble += v;
                        break;
                    case "-":
                        ValueDouble -= v;
                        break;
                    case "*":
                        ValueDouble *= v;
                        break;
                    case "/":
                        ValueDouble /= v;
                        break;
                    case "^":
                        ValueDouble = Math.Pow(ValueDouble, v);
                        break;
                }
            }
        }
        public class TokenResultRange : TokenResult
        {
            public TokenResultRange(Token t, FormulaRangeAddress v)
            {
                Token = t;
                Address = v;
            }
            public override void Negate()
            {

            }
            public override object Value
            {
                get
                {
                    return Address;
                }
            }
            public FormulaRangeAddress Address;
            public override void ApplyOperator(Token op, TokenResult tr)
            {
                if (op.Value == ":")
                {
                    if (tr.Token.TokenType == TokenType.CellAddress)
                    {
                        var a = ((TokenResultRange)tr).Address;
                        Address.FromRow = Address.FromRow < a.FromRow ? Address.FromRow : a.FromRow;
                        Address.ToRow = Address.ToRow > a.ToRow ? Address.ToRow : a.ToRow;
                        Address.FromCol = Address.FromCol < a.FromCol ? Address.FromCol : a.FromCol;
                        Address.ToCol = Address.ToCol > a.ToCol ? Address.ToCol : a.ToCol;
                    }
                }
            }
        }
        internal CompileResult Execute(List<Token> exps)
        {
            var cell = new RpnFormulaCell();
            short extRefIx = short.MinValue;
            short wsIx = short.MinValue;
            var s = cell._expressionStack;
            for (int i = 0; i < exps.Count; i++)
            {
                var t = exps[i];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                        s.Push(new RpnBooleanExpression(t.Value, _parsingContext));
                        break;
                    case TokenType.Integer:
                        s.Push(new RpnIntegerExpression(t.Value, _parsingContext));
                        break;
                    case TokenType.Decimal:
                        s.Push(new RpnDecimalExpression(t.Value, _parsingContext));
                        break;
                    case TokenType.StringContent:
                        s.Push(new RpnStringExpression(t.Value, _parsingContext));
                        break;                    
                    case TokenType.Negator:
                        s.Peek().Negate();
                        break;
                    case TokenType.CellAddress:
                        s.Push(new RpnCellAddressExpression(t.Value, _parsingContext, extRefIx, wsIx));
                        extRefIx = wsIx = short.MinValue;                        
                        break;
                    case TokenType.NameValue:
                        s.Push(new RpnNamedValueExpression(t.Value, _parsingContext, extRefIx, wsIx));
                        break;
                    case TokenType.ExternalReference:
                        extRefIx = short.Parse(t.Value);
                        break;
                    case TokenType.WorksheetNameContent:
                        wsIx = _parsingContext.Package.Workbook.Worksheets.GetPositionByToken(t.Value);
                        break;
                    case TokenType.Function:
                        ExecFunc(t, cell);
                        break;
                    case TokenType.StartFunctionArguments:
                        cell._funcStackPosition.Push(s.Count);
                        break;
                    case TokenType.OpeningEnumerable:
                        i = GetArray(exps, i, out List<List<object>> array);
                        s.Push(new RpnEnumerableExpression(array, _parsingContext);
                        break;
                    case TokenType.Operator:
                        ApplyOperator(t, cell);
                        break;
                }
            }
            return s.Pop().Compile();
        }

        private int GetArray(List<Token> exps, int i, out List<List<object>> matrix)
        {            
            matrix= new List<List<object>>();   
            var array = new List<object>();
            while (i++ < exps.Count && exps[i].TokenType != TokenType.Enumerable)
            {
                var t = exps[i];
                switch (t.TokenType)
                {
                    case TokenType.Boolean:
                        array.Add(bool.Parse(t.Value));
                        break;
                    case TokenType.Integer:
                        array.Add(int.Parse(t.Value));
                        break;
                    case TokenType.Decimal:
                        array.Add(double.Parse(t.Value));
                        break;
                    case TokenType.StringContent:
                        array.Add(t.Value);
                        break;
                    case TokenType.SemiColon:
                        matrix.Add(array);
                        array = new List<object>();
                        break;
                    case TokenType.Enumerable:
                        break;
                    default:
                        throw new InvalidFormulaException("Array contains invalid tokens. Cell "+ _parsingContext.CurrentCell.WorksheetIx);
                }
            }
            if(i==exps.Count || exps[i].TokenType != TokenType.Enumerable)
            {
                throw new InvalidFormulaException("Array is not closed. Cell " + _parsingContext.CurrentCell.WorksheetIx);
            }
            return i+1;
        }

        private void ExecFunc(Token t, RpnFormulaCell cell)
        {
            var f = _parsingContext.Configuration.FunctionRepository.GetFunction(t.Value);
            var args = GetFunctionArguments(cell);
            //var compilerFactory = new RpnFunctionCompilerFactory(_parsingContext.Configuration.FunctionRepository, _parsingContext);
            var compiler = _functionCompilerFactory.Create(f);
            var result = compiler.Compile(args);
            PushResult(cell, result);
        }

        private void PushResult(RpnFormulaCell cell, CompileResult result)
        {
            switch (result.DataType)
            {
                case DataType.Boolean:
                    cell._expressionStack.Push(new RpnBooleanExpression(result, _parsingContext));
                    break;
                case DataType.Integer:
                    cell._expressionStack.Push(new RpnDecimalExpression(result, _parsingContext));
                    break;
                case DataType.Decimal:
                    cell._expressionStack.Push(new RpnDecimalExpression(result, _parsingContext));
                    break;
                case DataType.String:
                    cell._expressionStack.Push(new RpnStringExpression(result, _parsingContext));
                    break;
                case DataType.ExcelRange:
                    cell._expressionStack.Push(new RpnRangeExpression((IRangeInfo)result.Result, _parsingContext));
                    break;
            }
        }

        private IList<RpnExpression> GetFunctionArguments(RpnFormulaCell cell)
        {
            var list = new List<RpnExpression>();
            var pos = cell._funcStackPosition.Pop();
            var s = cell._expressionStack;
            while (s.Count > pos)
            {
                var si = s.Pop();
                list.Insert(0, si);
            }
            return list;
        }

        private void ApplyOperator(Token opToken, RpnFormulaCell cell)
        {
            var v1 = cell._expressionStack.Pop();
            var v2 = cell._expressionStack.Pop();
            
            var c1 = v1.Compile();
            var c2 = v2.Compile();

            if (OperatorsDict.Instance.TryGetValue(opToken.Value, out IOperator op))
            {
                var result = op.Apply(c2, c1, _parsingContext);
                PushResult(cell, result);
            }
        }
    }
}
