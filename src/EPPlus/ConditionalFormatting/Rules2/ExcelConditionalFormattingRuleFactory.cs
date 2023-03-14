using OfficeOpenXml.Utils;
using System;
using System.Xml;

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

            switch (type)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:

                    return
                        new ExcelConditionalFormattingGreaterThan(
                          address,
                          priority,
                          worksheet);

                case eExcelConditionalFormattingRuleType.LessThan:
                    return new ExcelConditionalFormattingLessThan(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.Between:
                    return new ExcelConditionalFormattingBetween(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.Equal:
                    return new ExcelConditionalFormattingEqual(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return new ExcelConditionalFormattingContainsText(
                       address,
                       priority,
                       worksheet);

            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             type.ToString()));
        }

        public static ExcelConditionalFormattingRule Create(ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
        {

            string cfType = xr.GetAttribute("type");
            string op = xr.GetAttribute("operator");

            if (cfType == "cellIs")
            {
                cfType = op;
            }

            //Capitalize first letter
            cfType = cfType[0].ToString().ToUpper() + cfType.Substring(1);


            eExcelConditionalFormattingRuleType eType =
                (eExcelConditionalFormattingRuleType)
                    Enum.Parse(typeof(eExcelConditionalFormattingRuleType),
                        cfType);

            switch (eType)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    return new ExcelConditionalFormattingGreaterThan(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LessThan:
                    return new ExcelConditionalFormattingLessThan(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Between:
                    return new ExcelConditionalFormattingBetween(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Equal:
                    return new ExcelConditionalFormattingEqual(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return new ExcelConditionalFormattingContainsText(address, ws, xr);

            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             eType.ToString()));
        }
    }
}