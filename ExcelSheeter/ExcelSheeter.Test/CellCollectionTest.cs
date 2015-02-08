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
    public class CellCollectionTest
    {
        [TestMethod]
        public void CellCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new CellCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void CellCollection_ByKey_SetValue()
        {
            // Arrange
            var value = "value";
            var cell = new Cell(value);
            var list = new CellCollection();

            // Act
            list[0] = cell;

            // Assert
            Assert.AreEqual(1, list.Count);
        }

        [TestMethod]
        public void CellCollection_ByKey_GetValue()
        {
            // Arrange
            var value = "value";
            var cell = new Cell(value);

            var list = new CellCollection();
            list[0] = cell;

            // Act
            var returnedCell = list[0];

            // Assert
            Assert.AreEqual(cell, returnedCell);
        }
    }
}
