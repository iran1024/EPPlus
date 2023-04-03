﻿
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingContainsText : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingContainsText
    {
        public ExcelConditionalFormattingContainsText(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;
            ContainText = string.Empty;
        }

        public ExcelConditionalFormattingContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;

            Text = Formula.GetSubstringStoppingAtSymbol("NOT(ISERROR(SEARCH(\"".Length);
        }

        ExcelConditionalFormattingContainsText(ExcelConditionalFormattingContainsText copy) :base(copy)
        {
            ContainText = copy.ContainText;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingContainsText(this);
        }

        public string ContainText
        {
            get
            {
                return Text;
            }
            set
            {
                Text = value;
                //TODO: Error check/Throw when formula does not follow this format and is a ContainsText.
                Formula = string.Format(
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }
    }
}
