using OfficeOpenXml.Utils;
using System;
using System.Reflection;
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

            // According the conditional formatting rule type
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
            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             type.ToString()));

        }

        static Type GetRuleType(eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.AboveAverage:
                    return typeof(ExcelConditionalFormattingAboveAverage);

                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                    return typeof(ExcelConditionalFormattingAboveOrEqualAverage);

                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return typeof(ExcelConditionalFormattingBelowAverage);

                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                    return typeof(ExcelConditionalFormattingBelowOrEqualAverage);

                case eExcelConditionalFormattingRuleType.AboveStdDev:
                    return typeof(ExcelConditionalFormattingAboveStdDev);

                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return typeof(ExcelConditionalFormattingBelowStdDev);

                case eExcelConditionalFormattingRuleType.Bottom:
                    return typeof(ExcelConditionalFormattingBottom);

                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return typeof(ExcelConditionalFormattingBottomPercent);

                case eExcelConditionalFormattingRuleType.Top:
                    return typeof(ExcelConditionalFormattingTop);

                case eExcelConditionalFormattingRuleType.TopPercent:
                    return typeof(ExcelConditionalFormattingTopPercent);

                case eExcelConditionalFormattingRuleType.Last7Days:
                    return typeof(ExcelConditionalFormattingLast7Days);

                case eExcelConditionalFormattingRuleType.LastMonth:
                    return typeof(ExcelConditionalFormattingLastMonth);

                case eExcelConditionalFormattingRuleType.LastWeek:
                    return typeof(ExcelConditionalFormattingLastWeek);

                case eExcelConditionalFormattingRuleType.NextMonth:
                    return typeof(ExcelConditionalFormattingNextMonth);

                case eExcelConditionalFormattingRuleType.NextWeek:
                    return typeof(ExcelConditionalFormattingNextWeek);

                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return typeof(ExcelConditionalFormattingThisMonth);

                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return typeof(ExcelConditionalFormattingThisWeek);

                case eExcelConditionalFormattingRuleType.Today:
                    return typeof(ExcelConditionalFormattingToday);

                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return typeof(ExcelConditionalFormattingTomorrow);

                case eExcelConditionalFormattingRuleType.Yesterday:
                    return typeof(ExcelConditionalFormattingYesterday);

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return typeof(ExcelConditionalFormattingBeginsWith);

                case eExcelConditionalFormattingRuleType.Between:
                    return typeof(ExcelConditionalFormattingBetween);

                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return typeof(ExcelConditionalFormattingContainsBlanks);

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return typeof(ExcelConditionalFormattingContainsErrors);

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return typeof(ExcelConditionalFormattingContainsText);

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return typeof(ExcelConditionalFormattingDuplicateValues);

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return typeof(ExcelConditionalFormattingEndsWith);

                case eExcelConditionalFormattingRuleType.Equal:
                    return typeof(ExcelConditionalFormattingEqual);

                case eExcelConditionalFormattingRuleType.Expression:
                    return typeof(ExcelConditionalFormattingExpression);

                case eExcelConditionalFormattingRuleType.GreaterThan:
                    return typeof(ExcelConditionalFormattingGreaterThan);

                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    return typeof(ExcelConditionalFormattingGreaterThanOrEqual);

                case eExcelConditionalFormattingRuleType.LessThan:
                    return typeof(ExcelConditionalFormattingLessThan);

                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    return typeof(ExcelConditionalFormattingLessThanOrEqual);

                case eExcelConditionalFormattingRuleType.NotBetween:
                    return typeof(ExcelConditionalFormattingNotBetween);

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return typeof(ExcelConditionalFormattingNotContainsBlanks);

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return typeof(ExcelConditionalFormattingNotContainsErrors);

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return typeof(ExcelConditionalFormattingNotContainsText);

                case eExcelConditionalFormattingRuleType.NotEqual:
                    return typeof(ExcelConditionalFormattingNotEqual);

                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return typeof(ExcelConditionalFormattingUniqueValues);

                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    return typeof(ExcelConditionalFormattingThreeColorScale);

                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return typeof(ExcelConditionalFormattingTwoColorScale);
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    return typeof(ExcelConditionalFormattingThreeIconSet);
                case eExcelConditionalFormattingRuleType.FourIconSet:
                    return typeof(ExcelConditionalFormattingFourIconSet);
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return typeof(ExcelConditionalFormattingFiveIconSet);
                case eExcelConditionalFormattingRuleType.DataBar:
                    return typeof(ExcelConditionalFormattingDataBar);
            }
            return null;
        }

        public static ExcelConditionalFormattingRule Create(string address, XmlReader xr)
        {
            string cfType = xr.GetAttribute("type");
            string op = xr.GetAttribute("operator");

            if (cfType == "cellIs")
            {
                cfType = op;
            }

            //Capitalize first letter
            cfType = cfType[0].ToString().ToUpper() + cfType.Substring(1);

            eExcelConditionalFormattingRuleType eType = (eExcelConditionalFormattingRuleType)
                                                           Enum.Parse(typeof(eExcelConditionalFormattingRuleType),
                                                           cfType);

            Type type = GetRuleType(eType);

            if (type != null)
            {
                ConstructorInfo ctor = type.GetConstructor(new[] { typeof(string), typeof(XmlReader) });
                object instance = ctor.Invoke(new object[] { address, xr });
                return (ExcelConditionalFormattingRule)instance;
            }
            throw new Exception("Conditional Formatting Type Not Found");
        }
    }
}