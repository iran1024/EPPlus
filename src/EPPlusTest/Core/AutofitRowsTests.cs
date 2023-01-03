﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class AutofitRowsTests : TestBase
    {

        [TestMethod]
        public void AutofitRow_ShouldCalculateNewRowHeightWhenWrapTextIsTrue()
        {
            using (var pck = OpenPackage("AutofitRows_DefaultWidth_WrapText_True.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["A1"].Style.WrapText = true;
                sheet.Cells["A1"].AutoFitRows();
                Assert.AreEqual(122.2d, sheet.Row(1).Height);
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void AutofitRow_ShouldNotCalculateNewRowHeightWhenWrapTextIsTrue()
        {
            var defaultWidth = 15d;
            using (var pck = OpenPackage("AutofitRows_DefaultWidth_WrapText_False.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["A1"].Style.WrapText = false;
                sheet.Cells["A1"].AutoFitRows();
                Assert.AreEqual(defaultWidth, sheet.Row(1).Height);
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void AutofitRow_SetRowHeight_CustomWidth_Regular()
        {
            using (var pck = OpenPackage("AutofitRows_CustomWidth_WrapText_True_Regular.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["A1"].Style.WrapText = true;
                sheet.Cells["A1"].AutoFitRows();
                //Assert.AreEqual(118.3d, sheet.Row(1).Height);

                sheet.Column(2).Width = 25d;
                sheet.Cells["B2"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["B2"].Style.WrapText = true;
                sheet.Cells["B2"].AutoFitRows();
                //Assert.AreEqual(30.9d, sheet.Row(2).Height);

                sheet.Column(3).Width = 20d;
                sheet.Cells["C3"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["C3"].Style.WrapText = true;
                sheet.Cells["C3"].AutoFitRows();
                //Assert.AreEqual(45.5d, sheet.Row(3).Height);

                SaveAndCleanup(pck);
            }
        }

        [DataTestMethod]
        [DataRow("Calibri", true)]
        [DataRow("Arial", false)]
        [DataRow("Arial Black", false)]
        [DataRow("Times New Roman", false)]
        [DataRow("Courier New", false)]
        [DataRow("Liberation Serif", false)]
        [DataRow("Verdana", false)]
        [DataRow("Cambria", false)]
        [DataRow("Cambria Math", false)]
        [DataRow("Georgia", false)]
        [DataRow("Corbel", false)]
        [DataRow("Century Gothic", false)]
        [DataRow("Rockwell", false)]
        [DataRow("Trebuchet MS", false)]
        [DataRow("Tw Cen MT", false)]
        [DataRow("Tw Cen MT Condensed", false)]
        public void AutofitRow_SetRowHeight_CustomWidth_Regular_Linebreak(string fontName, bool delete)
        {
            using (var pck = OpenPackage("AutofitRows_CustomWidth_WrapText_True_Regular_Linebreak.xlsx", delete))
            {
                var sheet = pck.Workbook.Worksheets.Add(fontName);
                sheet.Cells.Style.Font.Name= fontName;

                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["A1"].Style.WrapText = true;
                sheet.Cells["A1"].AutoFitRows();

                sheet.Column(2).Width = 25d;
                sheet.Cells["B2"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["B2"].Style.WrapText = true;
                sheet.Cells["B2"].AutoFitRows();

                sheet.Column(3).Width = 20d;
                sheet.Cells["C3"].Value = "A long text that needs some serious autofit of row height\n\nAnd some more text that needs some autofit";
                sheet.Cells["C3"].Style.WrapText = true;
                sheet.Cells["C3"].AutoFitRows();

                pck.Settings.TextSettings.AutofitHeightScaleFactor = 0.9f;

                sheet.Cells["D4"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["D4"].Style.WrapText = true;
                sheet.Cells["D4"].AutoFitRows();

                sheet.Column(5).Width = 25d;
                sheet.Cells["E5"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["E5"].Style.WrapText = true;
                sheet.Cells["E5"].AutoFitRows();

                sheet.Column(6).Width = 20d;
                sheet.Cells["F6"].Value = "A long text that needs some serious autofit of row height\n\nAnd some more text that needs some autofit";
                sheet.Cells["F6"].Style.WrapText = true;
                sheet.Cells["F6"].AutoFitRows();

                SaveAndCleanup(pck);
            }
        }

        [DataTestMethod]
        [DataRow("Calibri")]
        [DataRow("Arial")]
        [DataRow("Arial Black")]
        [DataRow("Times New Roman")]
        [DataRow("Courier New")]
        [DataRow("Liberation Serif")]
        [DataRow("Verdana")]
        [DataRow("Cambria")]
        [DataRow("Cambria Math")]
        [DataRow("Georgia")]
        [DataRow("Corbel")]
        [DataRow("Century Gothic")]
        [DataRow("Rockwell")]
        [DataRow("Trebuchet MS")]
        [DataRow("Tw Cen MT")]
        [DataRow("Tw Cen MT Condensed")]
        public void AutofitRow_SetRowHeight_CustomWidth_Italic_Linebreak(string fontName)
        {
            using (var pck = OpenPackage("AutofitRows_CustomWidth_WrapText_True_Italic_Linebreak.xlsx", false))
            {
                var sheet = pck.Workbook.Worksheets.Add(fontName);
                sheet.Cells.Style.Font.Name = fontName;
                sheet.Cells.Style.Font.Italic = true;

                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["A1"].Style.WrapText = true;
                sheet.Cells["A1"].AutoFitRows();

                sheet.Column(2).Width = 25d;
                sheet.Cells["B2"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["B2"].Style.WrapText = true;
                sheet.Cells["B2"].AutoFitRows();

                sheet.Column(3).Width = 20d;
                sheet.Cells["C3"].Value = "A long text that needs some serious autofit of row height\n\nAnd some more text that needs some autofit";
                sheet.Cells["C3"].Style.WrapText = true;
                sheet.Cells["C3"].AutoFitRows();

                pck.Settings.TextSettings.AutofitHeightScaleFactor = 0.9f;

                sheet.Cells["D4"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["D4"].Style.WrapText = true;
                sheet.Cells["D4"].AutoFitRows();

                sheet.Column(5).Width = 25d;
                sheet.Cells["E5"].Value = "A long text that needs some serious autofit of row height\n\r\n\rAnd some more text that needs some autofit";
                sheet.Cells["E5"].Style.WrapText = true;
                sheet.Cells["E5"].AutoFitRows();

                sheet.Column(6).Width = 20d;
                sheet.Cells["F6"].Value = "A long text that needs some serious autofit of row height\n\nAnd some more text that needs some autofit";
                sheet.Cells["F6"].Style.WrapText = true;
                sheet.Cells["F6"].AutoFitRows();

                SaveAndCleanup(pck);
            }
        }
    }
}
