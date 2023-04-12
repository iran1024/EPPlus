using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingThreeIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>
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

    //    internal ExcelConditionalFormattingThreeIconSet(
    //    ExcelAddress address,
    //    int priority,
    //    ExcelWorksheet worksheet)
    //: base(
    //  eExcelConditionalFormattingRuleType.ThreeIconSet,
    //  address,
    //  priority,
    //  worksheet)
    //    {

    //    }
    }
}
