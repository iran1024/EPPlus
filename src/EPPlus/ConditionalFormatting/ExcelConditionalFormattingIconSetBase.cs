using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingIconSetBase<T> : 
        ExcelConditionalFormattingRule,
        IExcelConditionalFormattingThreeIconSet<T>
        where T : Enum
    {
        private string _uid = null;

        internal string Uid { 
            get 
            {
                if(_uid == null)
                {
                    return NewId();
                }

                return _uid;
            } 
            private set
            {
                _uid = value;
            }
        }

        internal static string NewId()
        {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        internal ExcelConditionalFormattingIconSetBase(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
            :base(type, 
                 address, 
                 priority, 
                 worksheet) 
        {
            double symbolCount;

            if (type == eExcelConditionalFormattingRuleType.ThreeIconSet)
            {
                symbolCount = 3;
            }
            else if(type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                symbolCount = 4;
            }
            else
            {
                symbolCount = 5;
            }

            Icon1 = CreateIcon(0, type);
            Icon2 = CreateIcon(Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(Math.Round(100D * (2D / symbolCount), 0), type);
        }

        protected ExcelConditionalFormattingIconDataBarValue CreateIcon(double value, eExcelConditionalFormattingRuleType type)
        {
            var icon = new ExcelConditionalFormattingIconDataBarValue
                (
                eExcelConditionalFormattingValueObjectType.Percent,
                type
                );

            icon.Value = value;

            return icon;
        }

        internal ExcelConditionalFormattingIconSetBase(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr, 
          ExcelConditionalFormattingRule rule)
            : base(rule)
        {
            var set = xr.GetAttribute("iconSet").Substring(1);

            Type = type;
            IconSet = set.ConvertToEnum<T>();

            double symbolCount;

            if (type == eExcelConditionalFormattingRuleType.ThreeIconSet)
            {
                symbolCount = 3;
            }
            else if (type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                symbolCount = 4;
            }
            else
            {
                symbolCount = 5;
            }

            Icon1 = CreateIcon(0, type);
            Icon2 = CreateIcon(Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(Math.Round(100D * (2D / symbolCount), 0), type);

            xr.Read();
            Icon1.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon1.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
            Icon2.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon2.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
            Icon3.Type = xr.GetAttribute("type").CapitalizeFirstLetter().ConvertToEnum<eExcelConditionalFormattingValueObjectType>();
            Icon3.Value = double.Parse(xr.GetAttribute("val"));

            xr.Read();
        }

        /// <summary>
        /// Settings for icon 1 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon1
        {
            get;
            internal set;
        }

        /// <summary>
        /// Settings for icon 2 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon2
        {
            get;
            internal set;
        }
        /// <summary>
        /// Settings for icon 2 in the iconset
        /// </summary>
        public ExcelConditionalFormattingIconDataBarValue Icon3
        {
            get;
            internal set;
        }

        /// <summary>
        /// Reverse the order of the icons
        /// </summary>
        public bool Reverse
        {
            get;
            set;
        }

        /// <summary>
        /// If the cell values are visible
        /// </summary>
        public bool ShowValue
        {
            get;
            set;
        }


        internal override bool IsExtLst
        {
            get
            {
                if (GetIconSetString() == "3Stars" ||
                    GetIconSetString() == "3Triangles" ||
                    GetIconSetString() == "5Boxes")
                {
                    return true;
                }

                return false;
            }
        }

        public T IconSet
        {
            get;
            set;
        }

        internal string GetIconSetString()
        {
            return GetIconSetString(IconSet);
        }

        internal string GetIconSetString(T value)
        {
            if (Type == eExcelConditionalFormattingRuleType.FourIconSet)
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "4Arrows";
                    case "ArrowsGray":
                        return "4ArrowsGray";
                    case "Rating":
                        return "4Rating";
                    case "RedToBlack":
                        return "4RedToBlack";
                    case "TrafficLights":
                        return "4TrafficLights";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
            else if (Type == eExcelConditionalFormattingRuleType.FiveIconSet)
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "5Arrows";
                    case "ArrowsGray":
                        return "5ArrowsGray";
                    case "Quarters":
                        return "5Quarters";
                    case "Rating":
                        return "5Rating";
                    case "Boxes":
                        return "5Boxes";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
            else
            {
                switch (value.ToString())
                {
                    case "Arrows":
                        return "3Arrows";
                    case "ArrowsGray":
                        return "3ArrowsGray";
                    case "Flags":
                        return "3Flags";
                    case "Signs":
                        return "3Signs";
                    case "Symbols":
                        return "3Symbols";
                    case "Symbols2":
                        return "3Symbols2";
                    case "TrafficLights1":
                        return "3TrafficLights1";
                    case "TrafficLights2":
                        return "3TrafficLights2";
                    case "Stars":
                        return "3Stars";
                    case "Triangles":
                        return "3Triangles";
                    default:
                        throw (new ArgumentException("Invalid type"));
                }
            }
        }
    }
}