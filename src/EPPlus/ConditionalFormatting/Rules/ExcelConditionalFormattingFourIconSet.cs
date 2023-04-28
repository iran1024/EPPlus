﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingFourIconSet : 
        ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting4IconsSetType>, 
        IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>
    {
        internal ExcelConditionalFormattingFourIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
            : base(
              eExcelConditionalFormattingRuleType.FourIconSet,
              address,
              priority,
              worksheet)
        {
            Icon4 = CreateIcon(75, eExcelConditionalFormattingRuleType.FourIconSet);
        }

        internal ExcelConditionalFormattingFourIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet,
        bool stopIfTrue,
        XmlReader xr)
            : base(
            eExcelConditionalFormattingRuleType.FourIconSet,
            address,
            priority,
            worksheet,
            stopIfTrue,
            xr)
        {
            Icon4 = CreateIcon(75, eExcelConditionalFormattingRuleType.FourIconSet);
            
            Icon4.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon4.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
            xr.Read();
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon4
        {
            get;
            internal set;
        }

        internal override bool IsExtLst
        {
            get
            {
                if (Icon1.CustomIcon != null ||
                    Icon2.CustomIcon != null ||
                    Icon3.CustomIcon != null ||
                    Icon4.CustomIcon != null)
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }
    }
}
