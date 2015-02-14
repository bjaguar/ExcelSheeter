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

namespace ExcelSheeter.Test
{
    [TestClass]
    public class WorksheetTest
    {
        [TestMethod]
        public void Worksheet_CreateNew()
        {
            // Arrange

            // Act
            var sheet = new Worksheet("sheet");

            // Assert
            Assert.IsNotNull(sheet);
            Assert.IsNotNull(sheet.Table);
            Assert.IsNotNull(sheet.Rows);
            Assert.IsNotNull(sheet.Columns);
            Assert.AreEqual("sheet", sheet.Name);
        }

        [TestMethod]
        public void Table_GetSet_Properties()
        {
            // Arrange
            var defaultColumnWidth = 10d;
            var defaultRowHeight = 11;
            var expandedColumnCount = 12;
            var expandedRowCount = 14;
            var leftCell = 15;
            var styleID = "style";
            var topCell = 16;
            var fullColumns = true;
            var fullRows = true;

            var sheet = new Worksheet("sheet");

            // Act
            sheet.DefaultColumnWidth = defaultColumnWidth;
            sheet.DefaultRowHeight = defaultRowHeight;
            sheet.ExpandedColumnCount = expandedColumnCount;
            sheet.ExpandedRowCount = expandedRowCount;
            sheet.LeftCell = leftCell;
            sheet.StyleId = styleID;
            sheet.TopCell = topCell;
            sheet.FullColumns = fullColumns;
            sheet.FullRows = fullRows;

            // Assert
            Assert.AreEqual(defaultColumnWidth, sheet.DefaultColumnWidth);
            Assert.AreEqual(defaultRowHeight, sheet.DefaultRowHeight);
            Assert.AreEqual(expandedColumnCount, sheet.ExpandedColumnCount);
            Assert.AreEqual(expandedRowCount, sheet.ExpandedRowCount);
            Assert.AreEqual(leftCell, sheet.LeftCell);
            Assert.AreEqual(styleID, sheet.StyleId);
            Assert.AreEqual(topCell, sheet.TopCell);
            Assert.AreEqual(fullColumns, sheet.FullColumns);
            Assert.AreEqual(fullRows, sheet.FullRows);
        }

        [TestMethod]
        public void Worksheet_GetCell_ByIndexes()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("sheet");
            sheet.Table.Rows.Add(row);

            // Act
            var returnedCell = sheet.GetCell(0, 1);

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Worksheet_GetCell_ByIndexAndName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("sheet");
            sheet.Table.Rows.Add(row);

            // Act
            var returnedCell = sheet.GetCell(0, "b");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Worksheet_GetCell_ByName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("sheet");
            sheet.Table.Rows.Add(row);

            // Act
            var returnedCell = sheet.GetCell("b1");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Worksheet_OuterXml()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("sheet");
            sheet.Table.Rows.Add(row);

            // Act

            // Assert
            Assert.IsNotNull(sheet.OuterXml);
        }

        [TestMethod]
        public void Worksheet_Get()
        {
            // Arrange
            var sheet = new Worksheet("name");

            // Act
            var cell = sheet["c5"];

            // Assert
            Assert.IsNotNull(cell);
        }
    }
}
