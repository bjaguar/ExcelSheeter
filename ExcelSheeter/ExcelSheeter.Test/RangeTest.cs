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
    public class RangeTest
    {
        [TestMethod]
        public void Range_CreateNew()
        {
            // Arrange

            // Act
            var range = new Range();

            // Assert
            Assert.IsNotNull(range);
            Assert.IsNotNull(range.Rows);
        }

        [TestMethod]
        public void Range_GetCell_ByIndexes()
        {
            // Arrange
            var range = new Range();

            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);
            
            range.Rows.Add(row);

            // Act
            var returnedCell = range.GetCell(0, 1);

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Range_GetCell_ByRowIndexAndColumnName()
        {
            // Arrange
            var range = new Range();

            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            range.Rows.Add(row);

            // Act
            var returnedCell = range.GetCell(0, "b");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Range_GetCell_ByName()
        {
            // Arrange
            var range = new Range();

            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            range.Rows.Add(row);

            // Act
            var returnedCell = range.GetCell("b1");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }
    }
}
