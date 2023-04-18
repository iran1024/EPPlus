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
            Icon4 = CreateIcon(address, worksheet, 75, eExcelConditionalFormattingRuleType.FourIconSet);
        }

        internal ExcelConditionalFormattingFourIconSet(
        ExcelAddress address,
        ExcelWorksheet worksheet,
        XmlReader xr,
        ExcelConditionalFormattingRule rule)
            : base(
            eExcelConditionalFormattingRuleType.FourIconSet,
            address,
            worksheet,
            xr,
            rule)
        {
            Icon4 = CreateIcon(address, worksheet, 75, eExcelConditionalFormattingRuleType.FourIconSet);
            
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
    }
}
