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
        public void Workbook_TryGetWorksheet()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.Sheets.Add(sheet);

            Worksheet returnedSheet;

            // Act
            var value = book.TryGetWorksheet("name", out returnedSheet);

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(sheet, returnedSheet);
        }

        [TestMethod]
        public void Workbook_TryGetWorksheet_WrongSheetName_ReturnsFalse()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.Sheets.Add(sheet);

            Worksheet returnedSheet;

            // Act
            var value = book.TryGetWorksheet("wrongname", out returnedSheet);

            // Assert
            Assert.IsFalse(value);
            Assert.IsNull(returnedSheet);
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
        public void Workbook_GetWorksheet_WrongSheetName_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.Sheets.Add(sheet);

            // Act
            try
            {
                var returnedSheet = book.GetWorksheet("wrongname");
            }
            catch { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void Workbook_AddWorksheet()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();

            // Act
            book.AddSheet(sheet);

            // Assert
            Assert.AreEqual(1, book.Sheets.Count);
        }

        [TestMethod]
        public void Workbook_AddWorksheet_ByName()
        {
            // Arrange
            var book = new Workbook();

            // Act
            book.AddSheet("name");

            // Assert
            Assert.AreEqual(1, book.Sheets.Count);
        }

        [TestMethod]
        public void Workbook_RemoveWorksheet()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.AddSheet(sheet);

            // Act
            var value = book.RemoveSheet(sheet);

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(0, book.Sheets.Count);
        }

        [TestMethod]
        public void Workbook_RemoveWorksheet_ByName()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.AddSheet(sheet);

            // Act
            var value = book.RemoveSheet("name");

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(0, book.Sheets.Count);
        }

        [TestMethod]
        public void Workbook_RemoveWorksheet_ByName_WrongName_ReturnsFalse()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.AddSheet(sheet);

            // Act
            var value = book.RemoveSheet("wrongname");

            // Assert
            Assert.IsFalse(value);
            Assert.AreEqual(1, book.Sheets.Count);
        }

        [TestMethod]
        public void Workbook_ClearWorksheets()
        {
            // Arrange
            var sheet = new Worksheet("name");

            var book = new Workbook();
            book.AddSheet(sheet);

            // Act
            book.ClearSheets();

            // Assert
            Assert.AreEqual(0, book.Sheets.Count);
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
            var returnedCell = sheet.GetCell(0, 1);

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
            var returnedCell = sheet.GetCell(0, "b");

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
            var returnedCell = sheet.GetCell("b1");

            // Assert
            Assert.AreEqual("value2", returnedCell.Value);
        }

        [TestMethod]
        public void Workbook_AddStyle()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();

            // Act
            book.AddStyle(style);

            // Assert
            Assert.AreEqual(1, book.Styles.Count);
        }

        [TestMethod]
        public void Workbook_AddStyle_ByName()
        {
            // Arrange
            var book = new Workbook();

            // Act
            var style = book.AddStyle("id");

            // Assert
            Assert.AreEqual(1, book.Styles.Count);
        }

        [TestMethod]
        public void Workbook_GetStyle()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            var returnedStyle = book.GetStyle("id");

            // Assert
            Assert.AreEqual(style, returnedStyle);
        }

        [TestMethod]
        public void Workbook_GetStyle_WrongId_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            try
            {
                var returnedStyle = book.GetStyle("wrongid");
            }
            catch { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void Workbook_TryGetStyle()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            Style returnedStyle;

            // Act
            var value = book.TryGetStyle("id", out returnedStyle);

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(style, returnedStyle);
        }

        [TestMethod]
        public void Workbook_TryGetStyle_WrongId_ReturnsFalse()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            Style returnedStyle;

            // Act
            var value = book.TryGetStyle("wrongid", out returnedStyle);

            // Assert
            Assert.IsFalse(value);
            Assert.IsNull(returnedStyle);
        }

        [TestMethod]
        public void Workbook_RemoveStyle()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            var value = book.RemoveStyle(style);

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(0, book.Styles.Count);
        }

        [TestMethod]
        public void Workbook_RemoveStyle_ById()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            var value = book.RemoveStyle("id");

            // Assert
            Assert.IsTrue(value);
            Assert.AreEqual(0, book.Styles.Count);
        }

        [TestMethod]
        public void Workbook_RemoveStyle_ByName_WrongId_ReturnsFalse()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            var value = book.RemoveStyle("wrongid");

            // Assert
            Assert.IsFalse(value);
            Assert.AreEqual(1, book.Styles.Count);
        }

        [TestMethod]
        public void Workbook_ClearStyle()
        {
            // Arrange
            var style = new Style("id");

            var book = new Workbook();
            book.AddStyle(style);

            // Act
            book.ClearStyles();

            // Assert
            Assert.AreEqual(0, book.Styles.Count);
        }
    }
}
