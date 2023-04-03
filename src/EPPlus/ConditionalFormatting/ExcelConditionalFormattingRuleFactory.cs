using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Utils;
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
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

                    return new ExcelConditionalFormattingGreaterThan(
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

                case eExcelConditionalFormattingRuleType.Yesterday:
                    return new ExcelConditionalFormattingYesterday(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Today:
                    return new ExcelConditionalFormattingToday(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return new ExcelConditionalFormattingTomorrow(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Last7Days:
                    return new ExcelConditionalFormattingLast7Days(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.LastWeek:
                    return new ExcelConditionalFormattingLastWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return new ExcelConditionalFormattingThisWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.NextWeek:
                    return new ExcelConditionalFormattingNextWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.LastMonth:
                    return new ExcelConditionalFormattingLastMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return new ExcelConditionalFormattingThisMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.NextMonth:
                    return new ExcelConditionalFormattingNextMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return new ExcelConditionalFormattingDuplicateValues(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return new ExcelConditionalFormattingTopBottomGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return new ExcelConditionalFormattingAverageGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return new ExcelConditionalFormattingStdDevGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.DataBar:
                    return new ExcelConditionalFormattingDataBar(
                        address, 
                        priority,
                        worksheet);
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return new ExcelConditionalFormattingTwoColorScale(
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

            if (cfType == "timePeriod")
            {
                cfType = xr.GetAttribute("timePeriod");
            }

            if (cfType == "top10")
            {
                bool isPercent = !string.IsNullOrEmpty(xr.GetAttribute("percent"));
                bool isBottom = !string.IsNullOrEmpty(xr.GetAttribute("bottom"));

                if (isPercent)
                {
                    cfType = "TopPercent";

                    if (isBottom)
                    {
                        cfType = "BottomPercent";
                    }
                }
                else if (isBottom)
                {
                    cfType = "Bottom";
                }
                else
                {
                    cfType = "Top";
                }
            }

            if (cfType == "aboveAverage")
            {
                //aboveAverage is true by default/when empty
                if (string.IsNullOrEmpty(xr.GetAttribute("aboveAverage")))
                {
                    cfType = "Above";
                }
                else
                {
                    cfType = "Below";
                }

                string stringEnding = "Average";

                if (!string.IsNullOrEmpty(xr.GetAttribute("stdDev")))
                {
                    stringEnding = "StdDev";
                }
                else if (!string.IsNullOrEmpty(xr.GetAttribute("equalAverage")))
                {
                    cfType = cfType + "OrEqual";
                }

                cfType = cfType + stringEnding;
            }

            string text = xr.GetAttribute("timePeriod");

            var eType = cfType.CapitalizeFirstLetter()
                        .ConvertToEnum<eExcelConditionalFormattingRuleType>();

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

                case eExcelConditionalFormattingRuleType.Yesterday:
                    return new ExcelConditionalFormattingYesterday(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Today:
                    return new ExcelConditionalFormattingToday(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return new ExcelConditionalFormattingTomorrow(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Last7Days:
                    return new ExcelConditionalFormattingLast7Days(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LastWeek:
                    return new ExcelConditionalFormattingLastWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return new ExcelConditionalFormattingThisWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NextWeek:
                    return new ExcelConditionalFormattingNextWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LastMonth:
                    return new ExcelConditionalFormattingLastMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return new ExcelConditionalFormattingThisMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NextMonth:
                    return new ExcelConditionalFormattingNextMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return new ExcelConditionalFormattingDuplicateValues(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return new ExcelConditionalFormattingTopBottomGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return new ExcelConditionalFormattingAverageGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return new ExcelConditionalFormattingStdDevGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.DataBar:
                    return new ExcelConditionalFormattingDataBar(address, ws, xr);

                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return new ExcelConditionalFormattingTwoColorScale(address, ws, xr);
            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             eType.ToString()));
        }
    }
}