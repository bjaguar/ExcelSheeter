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
    public class WorkbookTest
    {
        [TestMethod]
        public void Workbook_CreateNew()
        {
            // Arrange

            // Act
            var book = new Workbook();

            // Assert
            Assert.IsNotNull(book);
        }

        [TestMethod]
        public void Workbook_GetWorksheet()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            var returnedSheet = book.GetWorksheet("name");

            // Assert
            Assert.AreEqual(sheet, returnedSheet);
        }

        [TestMethod]
        public void Workbook_OuterXml()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            var xml = book.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void Workbook_GetCell_ByIndexes()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("name");
            sheet.Rows.Add(row);

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            var returnedCell = book.GetCell("name", 0, 1);

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Workbook_GetCell_ByIndexAndName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("name");
            sheet.Rows.Add(row);

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            var returnedCell = book.GetCell("name", 0, "b");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Workbook_GetCell_ByName()
        {
            // Arrange
            var cell1 = new Cell("value1");
            var cell2 = new Cell("value2");
            var cell3 = new Cell("value3");

            var row = new Row();
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);

            var sheet = new Worksheet("name");
            sheet.Rows.Add(row);

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            var returnedCell = book.GetCell("name", "b1");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }
    }
}
