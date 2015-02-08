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
    public class CellTest
    {
        [TestMethod]
        public void Cell_CreateNew()
        {
            // Arrange

            // Act
            var cell = new Cell();

            // Assert
            Assert.IsNotNull(cell);
        }

        [TestMethod]
        public void Cell_CreateNew_WithParams_Value()
        {
            // Arrange
            var value = "value";

            // Act
            var cell = new Cell(value);

            // Assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(value, cell.Value);
        }

        [TestMethod]
        public void Cell_CreateNew_WithParams_ValueAndDataType()
        {
            // Arrange
            var value = 1;
            var dataType = DataType.Number;

            // Act
            var cell = new Cell(value, dataType);

            // Assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(value, cell.Value);
            Assert.AreEqual(dataType, cell.DataType);
        }

        [TestMethod]
        public void Cell_SetValue()
        {
            // Arrange
            var value = "value";
            var newValue = "newValue";

            var cell = new Cell(value);

            // Act
            cell.Value = newValue;

            // Assert
            Assert.AreEqual(newValue, cell.Value);
        }

        [TestMethod]
        public void Cell_SetDataType()
        {
            // Arrange
            var value = 1;
            var dataType = DataType.String;
            var newDataType = DataType.Number;

            var cell = new Cell(value, dataType);

            // Act
            cell.DataType = newDataType;

            // Assert
            Assert.AreEqual(newDataType, cell.DataType);
        }

        [TestMethod]
        public void Cell_SetStyle()
        {
            // Arrange
            var style = "style";

            var cell = new Cell();

            // Act
            cell.Style = style;

            // Assert
            Assert.AreEqual(style, cell.Style);
        }

        [TestMethod]
        public void Cell_SetStyle_EmptyString_ClearsValue()
        {
            // Arrange
            var style = "style";

            var cell = new Cell();
            cell.Style = style;

            // Act
            cell.Style = string.Empty;

            // Assert
            Assert.AreEqual(0, cell.Attributes.Count);
        }

        [TestMethod]
        public void Cell_OuterXml()
        {
            // Arrange
            var value = 1;
            var dataType = DataType.Number;
            var cell = new Cell(value, dataType);

            // Act
            var xml = cell.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
