using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class ExcelConditionalFormattingThreeIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>
    {

    }

    public class ExcelConditionalFormattingIconSetBase<T>
    : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingThreeIconSet<T>
    {

    }
}
