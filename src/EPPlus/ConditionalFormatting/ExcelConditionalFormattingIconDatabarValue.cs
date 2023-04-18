using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// 18.3.1.11 cfvo (Conditional Format Value Object)
    /// Describes the values of the interpolation points in a gradient scale.
    /// </summary>
    public class ExcelConditionalFormattingIconDataBarValue
    {
        private eExcelConditionalFormattingRuleType _ruleType;
        private ExcelWorksheet _worksheet;

        internal bool HasValueOrFormula
        {
            get
            {
                if (Type != eExcelConditionalFormattingValueObjectType.Min
                    && Type != eExcelConditionalFormattingValueObjectType.Max)
                {
                    return true;
                }

                return false;
            }
        }

        //eExcelConditionalFormattingValueObjectType _valueType;

        public ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType valueType, eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            ExcelWorksheet worksheet)
        {
            RuleType = ruleType;
            Type = valueType;
            _worksheet = worksheet;
        }

        /// <summary>
        /// Rule type
        /// </summary>
        internal eExcelConditionalFormattingRuleType RuleType
        {
            get { return _ruleType; }
            set { _ruleType = value; }
        }

        eExcelConditionalFormattingValueObjectType _type;

        /// <summary>
        /// Value type
        /// </summary>
        public eExcelConditionalFormattingValueObjectType Type
        {
            get
            {
                return _type;
            }
            set
            {
                if ((_ruleType == eExcelConditionalFormattingRuleType.ThreeIconSet || _ruleType == eExcelConditionalFormattingRuleType.FourIconSet || _ruleType == eExcelConditionalFormattingRuleType.FiveIconSet) &&
                    (value == eExcelConditionalFormattingValueObjectType.Min || value == eExcelConditionalFormattingValueObjectType.Max))
                {
                    throw new ArgumentException("Value type can't be Min or Max for icon sets");
                }

                _type = value;
            }
        }

        /// <summary>
        /// Greater Than Or Equal 
        /// </summary>
        public bool GreaterThanOrEqualTo { get; set; }

        private double? _value = double.NaN;

        /// <summary>
        /// The value
        /// </summary>
        public double Value
        {
            get
            {
                if (Type == eExcelConditionalFormattingValueObjectType.Num
                    || Type == eExcelConditionalFormattingValueObjectType.Percent
                    || Type == eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    return (double)_value;
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                _value = null;

                // Only some types use the @val attribute
                if (Type == eExcelConditionalFormattingValueObjectType.Num
                    || Type == eExcelConditionalFormattingValueObjectType.Percent
                    || Type == eExcelConditionalFormattingValueObjectType.Percentile)
                {
                    _value = value;
                }
                else
                {
                    throw new InvalidOperationException("Value can only be changed if Type is Num, Percent or Percentile." +
                        $"Current Type is \"{Type}\"");
                }
            }
        }

        string _formula = "";

        /// <summary>
        /// The Formula of the Object Value (uses the same attribute as the Value)
        /// </summary>
        public string Formula
        {
            get
            {
                // Return empty if the Object Value type is not Formula
                if (Type != eExcelConditionalFormattingValueObjectType.Formula)
                {
                    return string.Empty;
                }

                // Excel stores the formula in the @val attribute
                return _formula;
            }
            set
            {
                // Only store the formula if the Object Value type is Formula
                if (Type == eExcelConditionalFormattingValueObjectType.Formula)
                {
                    _formula = value;
                }
                else
                {
                    throw new InvalidOperationException("Cannot store formula in a non-formula type");
                }
            }
        }
    }
}
