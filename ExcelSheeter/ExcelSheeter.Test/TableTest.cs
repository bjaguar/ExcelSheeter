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
    public class TableTest
    {
        [TestMethod]
        public void Table_CreateNew()
        {
            // Arrange

            // Act
            var table = new Table();

            // Assert
            Assert.IsNotNull(table);
            Assert.IsNotNull(table.Columns);
            Assert.IsNotNull(table.Rows);
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

            var table = new Table();

            // Act
            table.DefaultColumnWidth = defaultColumnWidth;
            table.DefaultRowHeight = defaultRowHeight;
            table.ExpandedColumnCount = expandedColumnCount;
            table.ExpandedRowCount = expandedRowCount;
            table.LeftCell = leftCell;
            table.StyleId = styleID;
            table.TopCell = topCell;
            table.FullColumns = fullColumns;
            table.FullRows = fullRows;

            // Assert
            Assert.AreEqual(defaultColumnWidth, table.DefaultColumnWidth);
            Assert.AreEqual(defaultRowHeight, table.DefaultRowHeight);
            Assert.AreEqual(expandedColumnCount, table.ExpandedColumnCount);
            Assert.AreEqual(expandedRowCount, table.ExpandedRowCount);
            Assert.AreEqual(leftCell, table.LeftCell);
            Assert.AreEqual(styleID, table.StyleId);
            Assert.AreEqual(topCell, table.TopCell);
            Assert.AreEqual(fullColumns, table.FullColumns);
            Assert.AreEqual(fullRows, table.FullRows);
        }

        [TestMethod]
        public void Table_GetSet_Properties_2()
        {
            // Arrange
            var fullColumns = false;
            var fullRows = false;

            var table = new Table();

            // Act
            table.FullColumns = fullColumns;
            table.FullRows = fullRows;

            // Assert
            Assert.AreEqual(fullColumns, table.FullColumns);
            Assert.AreEqual(fullRows, table.FullRows);
        }

        [TestMethod]
        public void Table_GetCell_ByIndexes()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var table = new Table();
            table.Rows.Add(row);

            // Act
            var returnedCell = table.GetCell(0, 1);

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Table_GetCell_ByIndexAndName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var table = new Table();
            table.Rows.Add(row);

            // Act
            var returnedCell = table.GetCell(0, "b");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Table_GetCell_ByName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var table = new Table();
            table.Rows.Add(row);

            // Act
            var returnedCell = table.GetCell("c1");

            // Assert
            Assert.AreEqual("value3", returnedCell.Value);
        }

        [TestMethod]
        public void Table_GetCell_ByName_NullName_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var table = new Table();
            table.Rows.Add(row);

            // Act
            try
            {
                var returnedCell = table.GetCell(null);
            }
            catch (ArgumentNullException) { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void Table_GetRange()
        {
            // Arrange
            var cell_a1 = new Cell("valuea1");
            var cell_a2 = new Cell("valuea2");
            var cell_a3 = new Cell("valuea3");
            var cell_b1 = new Cell("valueb1");
            var cell_b2 = new Cell("valueb2");
            var cell_b3 = new Cell("valueb3");

            var row1 = new Row();
            row1.Cells.Add(cell_a1);
            row1.Cells.Add(cell_a2);
            row1.Cells.Add(cell_a3);

            var row2 = new Row();
            row2.Cells.Add(cell_b1);
            row2.Cells.Add(cell_b2);
            row2.Cells.Add(cell_b3);
            
            var table = new Table();
            table.Rows.Add(row1);
            table.Rows.Add(row2);

            // Act
            var range = table.GetRange(1, 0, 1, 2);
            var returnedCell = range.GetCell(0, 0);

            // Assert
            Assert.IsNotNull(range);
            Assert.AreEqual(cell_b1, returnedCell);
        }

        [TestMethod]
        public void Table_OuterXml()
        {
            // Arrange
            var row = new Row();
            var column = new Column();

            var table = new Table();
            table.Rows.Add(row);
            table.Columns.Add(column);

            // Act
            var xml = table.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
