using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace ExcelSheeter.Integration.Test
{
    [TestClass]
    public class UnitTest2
    {
        [TestMethod]
        public void TestMethod1()
        {
            var workbook = new Workbook();

            var worksheet1 = workbook.AddSheet("sheet 1");
            for (int row = 1; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    var cell = worksheet1.GetCell(row, col);
                    cell.Value = string.Format("{0}:{1}", row + 1, col + 1);
                }
            }

            var worksheet2 = workbook.AddSheet("sheet 2");
            var style1 = workbook.AddStyle("style1");
            style1.AlignmentStyle.Horizontal = HorizontalAlignment.Center;
            style1.AlignmentStyle.ShrinkToFit = true;
            style1.Borders.Add(BorderStylePosition.Bottom, "#000000", BorderLineStyle.Continuous);
            style1.Borders.Add(BorderStylePosition.Left, "#000000", BorderLineStyle.Continuous);
            style1.Borders.Add(BorderStylePosition.Right, "#000000", BorderLineStyle.Continuous);
            style1.Borders.Add(BorderStylePosition.Top, "#000000", BorderLineStyle.Continuous);
            style1.FontStyle.Bold = true;
            style1.FontStyle.Family = FontFamily.Modern;
            style1.FontStyle.FontName = "Arial";
            style1.FontStyle.Size = 14;
            style1.FontStyle.Shadow = true;
            style1.FontStyle.Italic = true;
            style1.FontStyle.Color = "#ff0000";
            style1.InteriorStyle.Pattern = InteriorStylePattern.Solid;
            style1.InteriorStyle.Color = "#e0e0e0";
            var cellA1 = worksheet2["a1"];
            cellA1.Value = "value a1";
            cellA1.Style = "style1";
            var cellA2 = worksheet2["a2"];
            cellA2.Value = "value a2";
            cellA2.Style = "style1";
            var cellA3 = worksheet2["a3"];
            cellA3.Value = "value a3";
            cellA3.Style = "style1";
            var cellB1 = worksheet2["b1"];
            cellB1.Value = "value b1";
            cellB1.Style = "style1";
            var cellB2 = worksheet2["b2"];
            cellB2.Value = "value b2";
            cellB2.Style = "style1";
            var cellB3 = worksheet2["b3"];
            cellB3.Value = "value b3";
            cellB3.Style = "style1";

            // Assertions
            var outerXml = workbook.OuterXml;
            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            File.WriteAllText(path + @"\file2.xml", outerXml);
        }
    }
}
