using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal class ExcelConditionalFormattingCollection : IEnumerable<ExcelConditionalFormattingRule>
    {
        List<ExcelConditionalFormattingRule> _rules;
        ExcelWorksheet _ws;
        int LastPriority = 0;

        internal ExcelConditionalFormattingCollection(ExcelWorksheet ws)
        {
            _ws = ws;
        }

        internal ExcelConditionalFormattingCollection(XmlReader xr)
        {
        }

        IEnumerator<ExcelConditionalFormattingRule> IEnumerable<ExcelConditionalFormattingRule>.GetEnumerator()
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
            ExcelConditionalFormattingRule cfRule = ExcelConditionalFormattingRuleFactory.Create(
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
    }
}
