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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.ExternalReferences;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        private int _tokenIndex = 0;
        private FormulaAddressBase _currentAddress;
        private bool _negateNextExpression;
        private List<FormulaRangeAddress> _addresses;
        public ExpressionGraphBuilder(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
        {

        }

        public ExpressionGraphBuilder(IExpressionFactory expressionFactory, ParsingContext parsingContext)
        {
            _expressionFactory = expressionFactory;
            _parsingContext = parsingContext;
        }
        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            return Build(tokens, null);
        }
        public ExpressionGraph Build(IEnumerable<Token> tokens, List<FormulaRangeAddress> addresses)
        {
            _tokenIndex = 0;
            _graph.Reset();
            var tokensArr = tokens != null ? tokens.ToArray() : new Token[0];
            _addresses = addresses;
            BuildUp(tokensArr, null);
            return _graph;
        }

        private void BuildUp(Token[] tokens, Expression parent)
        {
            int bracketCount = 0;
            while (_tokenIndex < tokens.Length)
            {
                var token = tokens[_tokenIndex];
                IOperator op = null;
                if (token.TokenTypeIsSet(TokenType.OpeningBracket))
                {
                    bracketCount++;
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingBracket))
                {
                    bracketCount--;
                    if(bracketCount==0 && _currentAddress is FormulaTableAddress ta)
                    {
                        ta.SetTableAddress(_parsingContext.Package);
                        CreateAndAppendExpression(ref parent, ref token);
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.ExternalReference))
                {
                    _currentAddress = new FormulaCellAddress() { ExternalReferenceIx = (short)_parsingContext.Package.Workbook.ExternalLinks.GetExternalLink(token.Value) };
                }
                else if (token.TokenTypeIsSet(TokenType.WorksheetNameContent))
                {
                    if (_currentAddress == null)
                    {
                        _currentAddress = new FormulaCellAddress();
                    }
                    if(_currentAddress.ExternalReferenceIx == -1)
                    {
                        _currentAddress.WorksheetIx = _parsingContext.Package.Workbook.Worksheets.GetPositionByToken(token.Value);
                    }
                    else if(_currentAddress.ExternalReferenceIx > -1)
                    {
                        var er = _parsingContext.Package.Workbook.ExternalLinks[_currentAddress.ExternalReferenceIx];
                        if (er.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                        {
                            _currentAddress.WorksheetIx = (short)(((ExcelExternalWorkbook)er).CachedWorksheets[token.Value]?.SheetId ?? -1);
                        }
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.TableName))
                {
                    if(_currentAddress!=null)
                    {
                        _currentAddress = new FormulaTableAddress()
                        {
                            ExternalReferenceIx = _currentAddress.ExternalReferenceIx,
                            WorksheetIx = _currentAddress.WorksheetIx, 
                            TableName = token.Value 
                        };
                    }
                    else
                    {
                        _currentAddress = new FormulaTableAddress() { TableName = token.Value };
                    }
                }
                else if(token.TokenTypeIsSet(TokenType.TableColumn))
                {
                    var ta = (FormulaTableAddress)_currentAddress;
                    if (string.IsNullOrEmpty(ta.ColumnName1))
                    {
                        ta.ColumnName1 = token.Value;
                    }
                    else
                    {
                        ta.ColumnName2 = token.Value;
                    }
                }
                else if(token.TokenTypeIsSet(TokenType.TablePart))
                {
                    var ta = (FormulaTableAddress)_currentAddress;
                    if (string.IsNullOrEmpty(ta.TablePart1))
                    {
                        ta.TablePart1 = token.Value;
                    }
                    else
                    {
                        ta.TablePart2 = token.Value;
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.Operator) && OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    if(!(bracketCount > 0 && op.Operator==Operators.Colon) && !(_tokenIndex==0 && op==Operator.Eq))
                    {
                        SetOperatorOnExpression(parent, op);
                    }
                }
                //else if (token.TokenTypeIsSet(TokenType.RangeOffset))
                //{
                //    BuildRangeOffsetExpression(tokens, parent, token, tokenInfo);
                //}
                else if (token.TokenTypeIsSet(TokenType.Function))
                {                    
                    BuildFunctionExpression(tokens, parent, token.Value);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    _tokenIndex++;
                    BuildEnumerableExpression(tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    _tokenIndex++;
                    BuildGroupExpression(tokens, parent);
                }
                else if (token.TokenTypeIsSet(TokenType.ClosingParenthesis) || token.TokenTypeIsSet(TokenType.ClosingEnumerable))
                {
                    break;
                }
                else if (token.TokenTypeIsSet(TokenType.Negator))
                {
                    _negateNextExpression = !_negateNextExpression;
                }
                else if(token.TokenTypeIsSet(TokenType.Percent))
                {
                    SetOperatorOnExpression(parent, Operator.Percent);
                    if (parent == null)
                    {
                        _graph.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else if(!
                    (token.TokenTypeIsSet(TokenType.Comma) && bracketCount > 0 ||
                     token.TokenTypeIsSet(TokenType.WhiteSpace) || 
                     token.TokenTypeIsSet(TokenType.WorksheetName))
                    )
                {
                    CreateAndAppendExpression(ref parent, ref token);
                }
                _tokenIndex++;
            }
        }

        private void BuildEnumerableExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new EnumerableExpression(_parsingContext));
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                var enumerableExpression = new EnumerableExpression(_parsingContext);
                parent.AddChild(enumerableExpression);
                BuildUp(tokens, enumerableExpression);
            }
        }

        private void CreateAndAppendExpression(ref Expression parent, ref Token token)
        {
            if (IsWaste(token)) return;
            if (parent != null && 
                (token.TokenTypeIsSet(TokenType.Comma) || token.TokenTypeIsSet(TokenType.SemiColon)))
            {
                parent = parent.PrepareForNextChild();
                return;
            }
            if (_negateNextExpression)
            {
                token = token.CloneWithNegatedValue(true);
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token, _currentAddress);
            if (_currentAddress != null && _addresses!=null)
            {
                _addresses.Add((FormulaRangeAddress)_currentAddress);
            }

            _currentAddress = null;
            if (parent == null)
            {
                _graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenTypeIsSet(TokenType.String) || token.TokenTypeIsSet(TokenType.Colon))
            {
                return true;
            }
            return false;
        }

        //private void BuildRangeOffsetExpression(Token[] tokens, Expression parent, Token token, IDictionary<int, TokenInfo> tokenInfo)
        //{
        //    if(_nRangeOffsetTokens++ % 2 == 0)
        //    {
        //        _rangeOffsetExpression = new RangeOffsetExpression(_parsingContext);
        //        if(token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
        //        {
        //            _rangeOffsetExpression.OffsetExpression1 = new FunctionExpression("offset", _parsingContext, false);
        //            HandleFunctionArguments(tokens, _rangeOffsetExpression.OffsetExpression1, tokenInfo);
        //        }
        //        else if(token.TokenTypeIsSet(TokenType.ExcelAddress))
        //        {
        //            _rangeOffsetExpression.AddressExpression2 = _expressionFactory.Create(token) as ExcelAddressExpression;
        //        }
        //    }
        //    else
        //    {
        //        if (parent == null)
        //        {
        //            _graph.Add(_rangeOffsetExpression);
        //        }
        //        else
        //        {
        //            parent.AddChild(_rangeOffsetExpression);
        //        }
        //        if (token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset")
        //        {
        //            _rangeOffsetExpression.OffsetExpression2 = new FunctionExpression("offset", _parsingContext, _negateNextExpression);
        //            HandleFunctionArguments(tokens, _rangeOffsetExpression.OffsetExpression2, tokenInfo);
        //        }
        //        else if (token.TokenTypeIsSet(TokenType.ExcelAddress))
        //        {
        //            _rangeOffsetExpression.AddressExpression2 = _expressionFactory.Create(token) as ExcelAddressExpression;
        //        }
        //    }
        //}

        private void BuildFunctionExpression(Token[] tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                _graph.Add(new FunctionExpression(funcName, _parsingContext, _negateNextExpression));
                _negateNextExpression = false;
                HandleFunctionArguments(tokens, _graph.Current);
            }
            else
            {
                var func = new FunctionExpression(funcName, _parsingContext, _negateNextExpression);
                _negateNextExpression = false;
                parent.AddChild(func);
                HandleFunctionArguments(tokens, func);
            }
        }

        private void HandleFunctionArguments(Token[] tokens, Expression function)
        {
            _tokenIndex++;
            var token = tokens.ElementAt(_tokenIndex);
            if (!token.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                throw new ExcelErrorValueException(eErrorType.Value);
            }
            _tokenIndex++;
            BuildUp(tokens, function.Children.First());
        }

        private void BuildGroupExpression(Token[] tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new GroupExpression(_negateNextExpression, _parsingContext));
                _negateNextExpression = false;
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                if (parent.IsGroupedExpression || parent is FunctionArgumentExpression)
                {
                    var newGroupExpression = new GroupExpression(_negateNextExpression, _parsingContext);
                    _negateNextExpression = false;
                    parent.AddChild(newGroupExpression);
                    BuildUp(tokens, newGroupExpression);
                }
                 BuildUp(tokens, parent);
            }
        }

        private void SetOperatorOnExpression(Expression parent, IOperator op)
        {
            if (parent == null)
            {
                _graph.Current.Operator = op;
            }
            else
            {
                Expression candidate;
                if (parent is FunctionArgumentExpression)
                {
                    candidate = parent.Children.Last();
                }
                else
                {
                    candidate = parent.Children.Last();
                    if (candidate is FunctionArgumentExpression)
                    {
                        candidate = candidate.Children.Last();
                    }
                }
                candidate.Operator = op;
            }
        }
    }
}
