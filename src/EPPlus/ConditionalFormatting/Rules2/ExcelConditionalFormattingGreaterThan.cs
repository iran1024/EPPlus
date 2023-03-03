using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal class ExcelConditionalFormattingGreaterThan : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThan
    {

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingGreaterThan(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.GreaterThan, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThan;
            //Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingGreaterThan(string address, XmlReader xr) : base(address, xr)
        {
            Operator = string.IsNullOrEmpty(xr.GetAttribute("Operator")) ? null :
                (eExcelConditionalFormattingOperatorType)Enum.Parse(
                 typeof(eExcelConditionalFormattingOperatorType), xr.GetAttribute("Operator"));
        }

        #endregion Constructors

    }
}
