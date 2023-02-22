using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal static class ExcelConditionalFormattingRuleFactory
    {
        public static ExcelConditionalFormattingRule Create(
        eExcelConditionalFormattingRuleType type,
        ExcelAddress address,
        int priority, ExcelWorksheet worksheet)
        {
            Require.Argument(type);
            Require.Argument(address).IsNotNull("address");
            Require.Argument(worksheet).IsNotNull("worksheet");

            // According the conditional formatting rule type
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    return new ExcelConditionalFormattingGreaterThan(
                      address,
                      priority,
                      worksheet);
            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             type.ToString()));

        }
    }
}