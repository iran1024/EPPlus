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

            Icon1 = CreateIcon(address, worksheet, 0, type);
            Icon2 = CreateIcon(address, worksheet, Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(address, worksheet, Math.Round(100D * (2D / symbolCount), 0), type);
        }

        protected ExcelConditionalFormattingIconDataBarValue CreateIcon(ExcelAddress address, ExcelWorksheet worksheet, double value, eExcelConditionalFormattingRuleType type)
        {
            var icon = new ExcelConditionalFormattingIconDataBarValue
                (
                eExcelConditionalFormattingValueObjectType.Percent,
                type,
                address,
                worksheet
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

            Icon1 = CreateIcon(address, worksheet, 0, type);
            Icon2 = CreateIcon(address, worksheet, Math.Round(100D / symbolCount, 0), type);
            Icon3 = CreateIcon(address, worksheet, Math.Round(100D * (2D / symbolCount), 0), type);

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

        readonly Dictionary<int, string> _iconStringSetDictionary = new Dictionary<int, string>
            {
             { 0,  "3Arrows" },  
             { 1,  "3ArrowsGray" },
             { 2,  "3Flags" },
             { 3,  "3TrafficLights1" } ,
             { 4,  "3TrafficLights2" },
             { 5,  "3Signs" },
             { 6,  "3Symbols" },
             { 7,  "3Symbols2" },
             { 8,  "3Stars" },
             { 9,  "3Triangles" },
             { 10, "4Arrows" },
             { 11, "4ArrowsGray" },
             { 12, "4RedToBlack" },
             { 13, "4Rating" },
             { 14, "4TrafficLights" },
             { 15, "5Rating" },
             { 16, "5Quarters" },
             { 17, "5Boxes" },
             { 18, "NoIcons"},
            };

        //<KeyValuePair<Func<int, bool>, Action>>

        //readonly List<KeyValuePair<Func<int, bool>, string>> CustomIconNames = new List<KeyValuePair<Func<int, bool>, string>>
        //{
        //    {new KeyValuePair< x < 10, "3Arrows">>},

        //};

        internal virtual string GetCustomIconStringValue(ExcelConditionalFormattingIconDataBarValue icon)
        {
            if (icon.CustomIcon != null)
            {
                int customIconId = (int)icon.CustomIcon;

                var iconSetId = customIconId >> 4;

                return _iconStringSetDictionary[iconSetId];
            }

            throw new NotImplementedException($"Cannot get custom icon {icon} of {this} ");
        }

        internal int GetCustomIconIndex(ExcelConditionalFormattingIconDataBarValue icon)
        {
            if (icon.CustomIcon != null)
            {
                return (int)icon.CustomIcon & 0xf;
            }

            return -1;
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

        public T _iconSet;

        public T IconSet
        {
            get
            {
                ////var v = GetXmlNodeString(_iconSetPath);
                ////v = v.Substring(1); //Skip first icon.
                //return (T)Enum.Parse(typeof(T), v, true);
                return _iconSet;
            }
            set
            {
                _iconSet = value;
            }
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