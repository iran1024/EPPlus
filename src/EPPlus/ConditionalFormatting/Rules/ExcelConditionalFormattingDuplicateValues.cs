﻿using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingDuplicateValues : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingDuplicateValues
    {
        internal ExcelConditionalFormattingDuplicateValues(
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(eExcelConditionalFormattingRuleType.DuplicateValues, address, priority, worksheet)
        {

        }

        internal ExcelConditionalFormattingDuplicateValues(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(address, eExcelConditionalFormattingRuleType.DuplicateValues, ws, xr)
        {
        }
    }
}
