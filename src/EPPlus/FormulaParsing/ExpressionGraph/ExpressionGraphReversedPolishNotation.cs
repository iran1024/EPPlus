using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;


namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ExpressionGraphReversedPolishNotation
    {
        //https://www.andreinc.net/2010/10/05/converting-infix-to-rpn-shunting-yard-algorithm
        //https://en.wikipedia.org/wiki/Shunting_yard_algorithm#The_algorithm_in_detail
        public List<Token> CreateExpressionList(IList<Token> tokens)
        {
            var operators = OperatorsDict.Instance;
            Stack<Token> stack = new Stack<Token>();
            var expressions=new List<Token>();
            for(int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                switch(token.TokenType)
                {
                    case TokenType.OpeningParenthesis:
                        stack.Push(token);
                        break;
                    case TokenType.ClosingParenthesis:
                        if (stack.Count > 0)
                        {
                            var o = stack.Pop();
                            while(o.TokenType != TokenType.OpeningParenthesis)
                            {
                                expressions.Add(o);
                                if (stack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = stack.Pop();
                            }
                            if(stack.Count>0 && stack.Peek().TokenType == TokenType.Function)
                            {
                                expressions.Add(stack.Pop());
                            }
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if(stack.Count > 0)
                        {
                            var o2 = stack.Peek();
                            while (o2.TokenType == TokenType.Operator && 
                                (operators[o2.Value].Precedence < operators[token.Value].Precedence || 
                                (operators[o2.Value].Precedence == operators[token.Value].Precedence && token.Value!="^"))
                                ) 
                            {
                                expressions.Add(stack.Pop());
                                o2 = stack.Peek();
                            }
                        }
                        stack.Push(token);
                        break;
                    
                    case TokenType.Function:
                        stack.Push(token);
                        break;
                    case TokenType.Comma:
                    case TokenType.SemiColon:
                        break;
                    default:
                        expressions.Add(token);
                        break;
                }

            }

            while (stack.Count > 0)
            {
                expressions.Add(stack.Pop());
            }

            return expressions;
        }
    }
}
