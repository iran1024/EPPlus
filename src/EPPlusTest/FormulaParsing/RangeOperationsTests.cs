using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class RangeOperationsTests
    {
        [TestMethod]
        public void IntersectOperatorWithMultipleRanges()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Formula = "SUM(A1:A3 A2:A4 OFFSET(A1, 1, 0))";
                sheet.Calculate();
                var result = sheet.Cells["A5"].Value;
                Assert.AreEqual(2d, result);
            }
        }

        [TestMethod]
        public void AdditionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 + B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(6d, result);
            }
        }

        [TestMethod]
        public void SubtractionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 - B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(1d, result);
            }
        }

        [TestMethod]
        public void MultiplicationOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "SUM(A1:A2 * B1:B2)";
                sheet.Calculate();
                var result = sheet.Cells["B3"].Value;
                Assert.AreEqual(12d, result);
            }
        }

        [TestMethod]
        public void DivisionOperatorShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Formula = "AVERAGE(A1:A2 / B1:B2)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["B3"].Value, 2);
                Assert.AreEqual(1.75d, result);
            }
        }

        
    }
}
