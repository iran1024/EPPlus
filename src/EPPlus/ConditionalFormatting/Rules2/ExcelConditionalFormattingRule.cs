using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{

    internal class ExcelConditionalFormattingRule : IExcelConditionalFormattingRule
    {
        //Deprecated
        public XmlNode Node { get; }

        //Dictionary<int, string> rules = new Dictionary<int, string>
        //{
        //    { 0,"conditionalFormatting" },
        //    { 1,"sqref" },
        //    { 3,"cfRule" },
        //    { 4,"priority" },
        //    {5, }

        //};

        public eExcelConditionalFormattingRuleType Type { get; set; }
        public ExcelAddress Address { get; set; }
        public int Priority { get; set; }
        public bool StopIfTrue { get; set; }
        public bool PivotTable { get; set; }
        public ExcelDxfStyleConditionalFormatting Style { get; set; }

        private UInt16 _stdDev;

        public UInt16 StdDev
        {
            get
            {
                return _stdDev;
            }
            set
            {
                _stdDev = value == 0 ? (UInt16)1 : value;
            }
        }

        private UInt16 _rank;

        /// <summary>
        /// Rank (zero is not allowed and will be converted to 1)
        /// </summary>
        public UInt16 Rank
        {
            get
            {
                return _rank;
            }
            set
            {
                _rank = value == 0 ? (UInt16)1 : value;
            }
        }

        private ExcelWorksheet _ws;

        private int _DxfId = -1;

        /// <summary>
        /// The DxfId (Differential Formatting style id)
        /// </summary>
        internal int DxfId
        {
            get
            {
                return _DxfId;
            }
            set
            {
                _DxfId = value;
            }
        }

        #region Constructors
        /// <summary>
        /// Initalize <see cref="ExcelConditionalFormattingRule"/> from file
        /// </summary>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingRule(XmlReader xr)
        {

        }

        /// <summary>
        /// Initalize <see cref="ExcelConditionalFormattingRule"/> from variables
        /// </summary>
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
        {
            Require.That(address).IsNotNull();
            Require.That(worksheet).IsNotNull();

            _ws = worksheet;

            //string.Format()
            //move writing of root node.

            Address = address;
            Priority = priority;
            Type = type;

        }
        #endregion Constructors

        bool? _aboveAverage = null;
        string _aboveAverageValue;

        /// <summary>
        /// Above average
        /// </summary>
        internal protected bool? AboveAverage
        {
            get
            {
                // Above Average if TRUE or if attribute does not exists
                return (_aboveAverage == true) || (_aboveAverage == null);
            }
            set
            {
                // Only the types that needs the @AboveAverage
                if ((Type == eExcelConditionalFormattingRuleType.BelowAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowStdDev))
                {
                    _aboveAverageValue = "0";
                }
                _aboveAverage = value;
            }
        }

        bool? _equalAverage;

        /// <summary>
        /// EqualAverage
        /// </summary>
        internal protected bool? EqualAverage
        {
            get
            {
                // Equal Avarege only if TRUE
                return (_equalAverage == true);
            }
            set
            {
                string equalAverageValue = string.Empty;

                // Only the types that needs the @EqualAverage
                if ((Type == eExcelConditionalFormattingRuleType.AboveOrEqualAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage))
                {
                    equalAverageValue = "1";
                }

                _equalAverage = value;
            }
        }

        bool? _bottom;

        /// <summary>
        /// Bottom attribute
        /// </summary>
        internal protected bool? Bottom
        {
            get
            {
                // Bottom if TRUE
                return (_bottom == true);
            }
            set
            {
                string bottomValue = string.Empty;

                // Only the types that needs the @Bottom
                if ((Type == eExcelConditionalFormattingRuleType.Bottom)
                  || (Type == eExcelConditionalFormattingRuleType.BottomPercent))
                {
                    bottomValue = "1";
                }

                _bottom = value;
            }
        }

        bool? _percent;
        string _percentValue;

        /// <summary>
        /// Percent attribute
        /// </summary>
        internal protected bool? Percent
        {
            get
            {
                // Bottom if TRUE
                return (_percent == true);
            }
            set
            {
                string percentValue = string.Empty;

                // Only the types that needs the @Bottom
                if ((Type == eExcelConditionalFormattingRuleType.BottomPercent)
                  || (Type == eExcelConditionalFormattingRuleType.TopPercent))
                {
                    percentValue = "1";
                }

                _percentValue = percentValue;
                _percent = true;
            }
        }

        /// <summary>
        /// TimePeriod
        /// </summary>
        internal protected eExcelConditionalFormattingTimePeriodType TimePeriod { get; set; }

        /// <summary>
        /// Operator
        /// </summary>
        internal protected eExcelConditionalFormattingOperatorType Operator { get; set; }

        /// <summary>
        /// Formula
        /// </summary>
        public string Formula { get; set; }


        /// <summary>
        /// Formula2
        /// Note, no longer Requires Formula to be set before it.
        /// But will still throw error if both formulas not filled at save time.
        /// </summary>
        public string Formula2 { get; set; }

        private ExcelConditionalFormattingAsType _as = null;
        /// <summary>
        /// Provides access to type conversion for all conditional formatting rules.
        /// </summary>
        public ExcelConditionalFormattingAsType As
        {
            get
            {
                if (_as == null)
                {
                    _as = new ExcelConditionalFormattingAsType(this);
                }
                return _as;
            }
        }

        internal void SetStyle(ExcelDxfStyleConditionalFormatting style)
        {
            Style = style;
            DxfId = int.MinValue;
        }
    }
}
