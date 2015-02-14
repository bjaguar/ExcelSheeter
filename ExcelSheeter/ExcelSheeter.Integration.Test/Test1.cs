/*
 * Copyright 2015 Fernando Suárez Llamas
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 */

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace ExcelSheeter.Integration.Test
{
    [TestClass]
    public class Test1
    {
        private readonly string[][] Items = new string[][]
        {
            new string[] { "row 1", "name 1", "surname 1", "998877551", "address 1", "21", DateTime.Now.ToShortDateString() },
            new string[] { "row 2", "name 2", "surname 2", "998877552", "address 2", "22", DateTime.Now.ToShortDateString() },
            new string[] { "row 3", "name 3", "surname 3", "998877553", "address 3", "23", DateTime.Now.ToShortDateString() },
            new string[] { "row 4", "name 4", "surname 4", "998877554", "address 4", "24", DateTime.Now.ToShortDateString() },
            new string[] { "row 5", "name 5", "surname 5", "998877555", "address 5", "25", DateTime.Now.ToShortDateString() },
            new string[] { "row 6", "name 6", "surname 6", "998877556", "address 6", "26", DateTime.Now.ToShortDateString() },
            new string[] { "row 7", "name 7", "surname 7", "998877557", "address 7", "27", DateTime.Now.ToShortDateString() },
        };

        private void CreateStyles(Workbook workbook)
        {
            var style1 = workbook.AddStyle("id1");
            style1.InteriorStyle.Pattern = InteriorStylePattern.Solid;
            style1.InteriorStyle.Color = "#ededed";


            var style2 = workbook.AddStyle("id2");
            style2.BordersStyle.Borders.Add(new BorderStyle(BorderStylePosition.Bottom, "black", BorderLineStyle.Continuous));
            style2.BordersStyle.Borders.Add(new BorderStyle(BorderStylePosition.Left, "black", BorderLineStyle.Continuous));
            style2.BordersStyle.Borders.Add(new BorderStyle(BorderStylePosition.Right, "black", BorderLineStyle.Continuous));
            style2.BordersStyle.Borders.Add(new BorderStyle(BorderStylePosition.Top, "black", BorderLineStyle.Continuous));
            style2.FontStyle.Bold = true;
            style2.FontStyle.Italic = true;
            style2.AlignmentStyle.Horizontal = HorizontalAlignment.Right;
        }

        private void CreateSheet(Worksheet worksheet)
        {
            // Add data
            foreach (var itemRow in Items)
            {
                var row = new Row();

                foreach (var itemCell in itemRow)
                {
                    var cell = new Cell(itemCell);

                    row.Cells.Add(cell);
                }

                worksheet.Rows.Add(row);
            }
        }

        [TestMethod]
        public void TestMethod1()
        {
            var workbook = new Workbook();
            CreateStyles(workbook);

            var worksheet = workbook.AddSheet("Sheet 1");
            CreateSheet(worksheet);

            foreach (var row in worksheet.Rows)
            {
                var aCell = row.GetCell("a");
                aCell.Style = "style1";

                var bCell = row.GetCell("b");
                bCell.Style = "style2";

                var fCell = row.GetCell("f");
                fCell.DataType = DataType.Number;

                var gCell = row.GetCell("g");
                //gCell.DataType = DataType.DateTime;
            }

            var outerXml = workbook.OuterXml;

            // Assertions
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            File.WriteAllText(path + @"\file.xml", outerXml);
        }
    }
}
