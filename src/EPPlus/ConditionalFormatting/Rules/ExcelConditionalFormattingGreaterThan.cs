﻿using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
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
        internal ExcelConditionalFormattingGreaterThan(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.GreaterThan, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThan;
        }

        #endregion Constructors
    }
}
