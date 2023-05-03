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
    internal class ExcelConditionalFormattingFiveIconSet : 
        ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>, 
        IExcelConditionalFormattingFiveIconSet
    {
        internal ExcelConditionalFormattingFiveIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
            : base(
              eExcelConditionalFormattingRuleType.FiveIconSet,
              address,
              priority,
              worksheet)
        {
            Icon4 = CreateIcon(60, eExcelConditionalFormattingRuleType.FiveIconSet);
            Icon5 = CreateIcon(80, eExcelConditionalFormattingRuleType.FiveIconSet);
        }

        internal ExcelConditionalFormattingFiveIconSet(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet,
            bool stopIfTrue,
            XmlReader xr)
            : base(
            eExcelConditionalFormattingRuleType.FiveIconSet,
            address,
            priority,
            worksheet,
            stopIfTrue,
            xr)
        {
            Icon4 = CreateIcon(60, eExcelConditionalFormattingRuleType.FiveIconSet);
            Icon5 = CreateIcon(80, eExcelConditionalFormattingRuleType.FiveIconSet);

            Icon4.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon4.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();

            Icon5.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon5.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
            xr.Read();
        }

        internal ExcelConditionalFormattingFiveIconSet(ExcelConditionalFormattingFiveIconSet copy) : base(copy)
        {
            Icon4 = copy.Icon4;
            Icon5 = copy.Icon5;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingFiveIconSet(this);
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon4
        {
            get;
            internal set;
        }

        /// <summary>
        /// Icon 4 value
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon5
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
                    Icon4.CustomIcon != null ||
                    Icon5.CustomIcon != null)
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }
    }
}
