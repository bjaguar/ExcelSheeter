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
    public class RowTest
    {
        [TestMethod]
        public void Row_CreateNew()
        {
            // Arrange

            // Act
            var row = new Row();

            // Assert
            Assert.IsNotNull(row);
        }

        [TestMethod]
        public void Row_GetCell_ByColumnName()
        {
            // Arrange
            var value1 = "value1";
            var value2 = "value2";
            var value3 = "value3";

            var row = new Row();
            row.Cells.Add(new Cell(value1));
            row.Cells.Add(new Cell(value2));
            row.Cells.Add(new Cell(value3));

            // Act
            var returnedCell1 = row.GetCell("a");
            var returnedCell2 = row.GetCell("b");
            var returnedCell3 = row.GetCell("c");

            // Assert
            Assert.AreEqual(value1, returnedCell1.Value);
            Assert.AreEqual(value2, returnedCell2.Value);
            Assert.AreEqual(value3, returnedCell3.Value);
        }

        [TestMethod]
        public void Row_GetCell_ByColumnName_EmptyName_ThrowsException()
        {
            // Arrange
            var paramName = string.Empty;
            var value1 = "value1";
            var value2 = "value2";
            var value3 = "value3";

            var row = new Row();
            row.Cells.Add(new Cell(value1));
            row.Cells.Add(new Cell(value2));
            row.Cells.Add(new Cell(value3));

            // Act
            try
            {
                var returnedCell1 = row.GetCell(string.Empty);
            }
            catch (ArgumentNullException ex) { paramName = ex.ParamName; }

            // Assert
            Assert.AreEqual("columnName", paramName);
        }

        [TestMethod]
        public void Row_GetCell_ByColumnName_TwoLetterColumns()
        {
            // Arrange
            var columnLetters = "abcdefghijklmnopqrstuvwxyz";

            var row = new Row();

            for (var i = 0; i < columnLetters.Length; i++)
            {
                var value = columnLetters[i].ToString();

                row.Cells.Add(new Cell(value));
            }

            for (var i = 0; i < columnLetters.Length; i++)
            {
                for (var j = 0; j < 1; j++)
                {
                    var value = string.Format("{0}{1}", columnLetters[j], columnLetters[i]);

                    row.Cells.Add(new Cell(value));
                }
            }

            // Act
            var returnedCell1 = row.GetCell("aa");
            var returnedCell2 = row.GetCell("ab");
            var returnedCell3 = row.GetCell("az");

            // Assert
            Assert.AreEqual("aa", returnedCell1.Value);
            Assert.AreEqual("ab", returnedCell2.Value);
            Assert.AreEqual("az", returnedCell3.Value);
        }

        [TestMethod]
        public void Row_GetCell_ByColumnIndex()
        {
            // Arrange
            var value1 = "value1";
            var value2 = "value2";
            var value3 = "value3";

            var row = new Row();
            row.Cells.Add(new Cell(value1));
            row.Cells.Add(new Cell(value2));
            row.Cells.Add(new Cell(value3));

            // Act
            var returnedCell1 = row.GetCell(0);
            var returnedCell2 = row.GetCell(1);
            var returnedCell3 = row.GetCell(2);

            // Assert
            Assert.AreEqual(value1, returnedCell1.Value);
            Assert.AreEqual(value2, returnedCell2.Value);
            Assert.AreEqual(value3, returnedCell3.Value);
        }

        [TestMethod]
        public void Row_GetAndSet_Caption()
        {
            // Arrange
            var value = "value";
            var row = new Row();

            // Act
            row.Caption = value;
            var returnedValue = row.Caption;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Caption_EmptyValue_RemovesAttribute()
        {
            // Arrange
            var value = "value";
            var row = new Row();
            row.Caption = value;

            // Act
            row.Caption = string.Empty;

            // Assert
            Assert.AreEqual(0, row.Attributes.Count);
        }

        [TestMethod]
        public void Row_GetAndSet_AutoFitWidth()
        {
            // Arrange
            var value = true;
            var row = new Row();

            // Act
            row.AutoFitWidth = value;
            var returnedValue = row.AutoFitWidth;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_AutoFitWidth_FalseValue()
        {
            // Arrange
            var value = false;
            var row = new Row();

            // Act
            row.AutoFitWidth = value;
            var returnedValue = row.AutoFitWidth;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Height()
        {
            // Arrange
            var value = 10d;
            var row = new Row();

            // Act
            row.Height = value;
            var returnedValue = row.Height;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Hidden_TrueValue()
        {
            // Arrange
            var value = true;
            var row = new Row();

            // Act
            row.Hidden = value;
            var returnedValue = row.Hidden;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Hidden_FalseValue()
        {
            // Arrange
            var value = false;
            var row = new Row();

            // Act
            row.Hidden = value;
            var returnedValue = row.Hidden;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Index()
        {
            // Arrange
            var value = 2;
            var row = new Row();

            // Act
            row.Index = value;
            var returnedValue = row.Index;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_Span()
        {
            // Arrange
            var value = 1;
            var row = new Row();

            // Act
            row.Span = value;
            var returnedValue = row.Span;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_StyleID()
        {
            // Arrange
            var value = "value";
            var row = new Row();

            // Act
            row.StyleId = value;
            var returnedValue = row.StyleId;

            // Assert
            Assert.AreEqual(value, returnedValue);
        }

        [TestMethod]
        public void Row_GetAndSet_StyleID_EmptyValue_RemovesAttribute()
        {
            // Arrange
            var value = "value";
            var row = new Row();
            row.StyleId = value;

            // Act
            row.StyleId = string.Empty;

            // Assert
            Assert.AreEqual(0, row.Attributes.Count);
        }

        [TestMethod]
        public void Row_OuterXml()
        {
            // Arrange
            var value1 = "value1";
            var value2 = "value2";
            var value3 = "value3";

            var row = new Row();
            row.Cells.Add(new Cell(value1));
            row.Cells.Add(new Cell(value2));
            row.Cells.Add(new Cell(value3));

            // Act
            var xml = row.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
