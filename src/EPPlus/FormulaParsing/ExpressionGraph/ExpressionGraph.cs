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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionTree
    {
        private List<Expression> _expressions = new List<Expression>();
        IList<Expression> _addressExpressions=null;
        public IList<Expression> AddressExpressions 
        { 
            get
            {
                if (_addressExpressions == null)
                {
                    _addressExpressions = new List<Expression>();
                    GetAddressExpressions(_addressExpressions, _expressions);
                }
                return _addressExpressions;
            }
        }

        private void GetAddressExpressions(IList<Expression> list, IList<Expression> expressions)
        {
            foreach(var e in expressions)
            {
                if(e.ExpressionType==ExpressionType.CellAddress || e.ExpressionType==ExpressionType.RangeAddress)
                {
                    var a=e.Compile().Address;
                    if(a!=null)
                    {
                        list.Add(e);
                    }
                }
                if(e.HasChildren)
                {
                    GetAddressExpressions(list, e.Children);
                }
            }
        }

        public IList<Expression> Expressions { get { return _expressions; } }
        public Expression Current { get; private set; }

        public Expression Add(Expression expression)
        {
            _expressions.Add(expression);
            Current = expression;
            return expression;
        }

        public void Reset()
        {
            _expressions.Clear();
            Current = null;           
            _addressExpressions=null;
        }

        public void Remove(Expression item)
        {
            _expressions.Remove(item);
        }

        internal void SetAddresses(int rowOffset, int colOffset)
        {
            foreach(CellAddressExpression a in AddressExpressions)
            {
                
            }
        }

        internal ExpressionTree CreateFromOffset(int rowOffset, int colOffset)
        {
            var ret = new ExpressionTree();
            foreach(var e in _expressions)
            {
                ret.Add(e.Clone(rowOffset, colOffset));
            }
            return ret;
        }
    }
}
