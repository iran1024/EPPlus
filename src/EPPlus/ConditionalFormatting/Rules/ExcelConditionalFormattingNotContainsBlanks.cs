﻿using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotContainsBlanks : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotContainsBlanks
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNotContainsBlanks(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsBlanks,
                address,
                priority,
                worksheet
                )
        {
            UpdateFormula();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingNotContainsBlanks(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsBlanks,
                address,
                worksheet,
                xr)
        {
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }

        void UpdateFormula()
        {
            Formula = string.Format(
              "LEN(TRIM({0}))>0",
              Address.Start.Address);
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
