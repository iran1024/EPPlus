using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingThreeIconSet : 
        ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>
    {
        internal ExcelConditionalFormattingThreeIconSet(
        ExcelAddress address,
        int priority,
        ExcelWorksheet worksheet)
            : base(
              eExcelConditionalFormattingRuleType.ThreeIconSet,
              address,
              priority,
              worksheet)
        {
        }

        internal ExcelConditionalFormattingThreeIconSet(
        ExcelAddress address,
        ExcelWorksheet worksheet,
        XmlReader xr,
        ExcelConditionalFormattingRule rule)
            : base(
            eExcelConditionalFormattingRuleType.ThreeIconSet,
            address,
            worksheet,
            xr,
            rule)
        {
            xr.Read();
        }

        internal override bool IsExtLst
        {
            get
            {
                if ( Icon1.CustomIcon != null ||
                     Icon2.CustomIcon != null ||
                     Icon3.CustomIcon != null )
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }
    }
}
