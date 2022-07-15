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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public abstract class Expression
    {        
        internal string ExpressionString { get; private set; }
        private readonly List<Expression> _children = new List<Expression>();
        protected ParsingContext Context { get; private set; }
        public IList<Expression> Children { get { return _children; } }
        public Expression Next { get; set; }
        public Expression Prev { get; set; }
        public IOperator Operator { get; set; }
        internal abstract ExpressionType ExpressionType { get; }
        public abstract bool IsGroupedExpression { get; }
        /// <summary>
        /// If set to true, <see cref="ExcelAddressExpression"></see>s that has a circular reference to their cell will be ignored when compiled
        /// </summary>
        public virtual bool IgnoreCircularReference
        {
            get; set;
        }

        public Expression(ParsingContext ctx)
        {
            Context = ctx;
        }

        public Expression(string expression, ParsingContext ctx)
        {
            ExpressionString = expression;
            Operator = null;
            Context = ctx;
        }

        public virtual bool HasChildren
        {
            get { return _children.Any(); }
        }

        /// <summary>
        /// Prepares the expression for next child expression
        /// </summary>
        /// <returns></returns>
        public virtual Expression  PrepareForNextChild()
        {
            return this;
        }

        /// <summary>
        /// Prepares the expression for next child expression.
        /// </summary>
        /// <param name="token"><see cref="Token"/> that is relevant in the context.</param>
        /// <returns></returns>
        public virtual Expression PrepareForNextChild(Token token)
        {
            return this;
        }

        /// <summary>
        /// Adds a child expression.
        /// </summary>
        /// <param name="child">The child expression to add</param>
        /// <returns></returns>
        public virtual Expression AddChild(Expression child)
        {
            if (_children.Any())
            {
                var last = _children.Last();
                child.Prev = last;
                last.Next = child;
            }
            _children.Add(child);
            return child;
        }

        public virtual Expression MergeWithNext()
        {
            var expression = this;
            if (Next != null && Operator != null)
            {
                var left = Compile();
                var right = Next.Compile();
                var result = Operator.Apply(left, right, Context);
                expression = ExpressionConverter.GetInstance(Context).FromCompileResult(result);
                if (expression is ExcelErrorExpression)
                {
                    expression.Next = null;
                    expression.Prev = null;
                    return expression;
                }
                if (Next != null)
                {
                    expression.Operator = Next.Operator;
                }
                else
                {
                    expression.Operator = null;
                }
                expression.Next = Next.Next;
                if (expression.Next != null) expression.Next.Prev = expression;
                expression.Prev = Prev;
            }
            else
            {
                throw (new FormatException("Invalid formula syntax. Operator missing expression."));
            }
            if (Prev != null)
            {
                Prev.Next = expression;
            }            
            return expression;
        }
        protected CompileResult _result;
        public abstract CompileResult Compile();

    }
}
