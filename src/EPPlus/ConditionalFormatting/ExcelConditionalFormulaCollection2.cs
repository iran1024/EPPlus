using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingCollection2 : IEnumerable<Rules2.ExcelConditionalFormattingRule>
    {
        List<Rules2.ExcelConditionalFormattingRule> _rules = new List<Rules2.ExcelConditionalFormattingRule>();
        ExcelWorksheet _ws;
        int LastPriority = 1;

        internal ExcelConditionalFormattingCollection2(ExcelWorksheet ws)
        {
            _ws = ws;
            _rules = new List<Rules2.ExcelConditionalFormattingRule>();
        }

        internal ExcelConditionalFormattingCollection2(XmlReader xr, ExcelWorksheet ws)
        {
            _ws = ws;

            while (xr.ReadUntil(1, "conditionalFormatting", "sheetData", "dataValidations", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks", "extLst", "pageMargins"))
            {
                if(xr.LocalName == "conditionalFormatting")
                {
                    string address = xr.GetAttribute("sqref");

                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        xr.Read();
                        var cf = Rules2.ExcelConditionalFormattingRuleFactory.Create(new ExcelAddress(address), _ws, xr);

                        _rules.Add(cf);
                    }

                    xr.Read();
                    xr.Read();
                }
            }
        }



        IEnumerator<Rules2.ExcelConditionalFormattingRule> IEnumerable<Rules2.ExcelConditionalFormattingRule>.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        /// <summary>
        /// Add rule (internal)
        /// </summary>
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <returns></returns>F
        internal IExcelConditionalFormattingRule AddRule(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address)
        {
            Require.Argument(address).IsNotNull("address");

            // address = ValidateAddress(address);

            // Create the Rule according to the correct type, address and priority
            Rules2.ExcelConditionalFormattingRule cfRule = Rules2.ExcelConditionalFormattingRuleFactory.Create(
              type,
              address,
              LastPriority++,
              _ws);

            // Add the newly created rule to the list
            _rules.Add(cfRule);

            // Return the newly created rule
            return cfRule;
        }

        /// <summary>
        /// Add GreaterThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingGreaterThan)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThan,
              address);
        }

        public IExcelConditionalFormattingLessThan AddLessThan(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingLessThan)AddRule(
              eExcelConditionalFormattingRuleType.LessThan,
              address);
        }

        public IExcelConditionalFormattingBetween AddBetween(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingBetween)AddRule(
              eExcelConditionalFormattingRuleType.Between,
              address);
        }

        public IExcelConditionalFormattingEqual AddEqual(ExcelAddress address)
        {
            return (IExcelConditionalFormattingEqual)AddRule(
              eExcelConditionalFormattingRuleType.Equal,
              address);
        }

        public IExcelConditionalFormattingContainsText AddTextContains(ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddYesterday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Yesterday,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddToday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Today,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Tomorrow,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Last7Days,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastWeek,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisWeek,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextWeek,
              address);
        }
    }
}
