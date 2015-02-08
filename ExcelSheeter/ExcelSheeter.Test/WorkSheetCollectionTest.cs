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
    public class WorksheetCollectionTest
    {
        [TestMethod]
        public void WorksheetCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new WorksheetCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void WorksheetCollection_Contains()
        {
            // Arrange
            var sheet = new Worksheet("sheet");

            var list = new WorksheetCollection();
            list.Add(sheet);

            // Act
            var value = list.Contains("sheet");

            // Assert
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void WorksheetCollection_AddDuplicateName_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var sheet1 = new Worksheet("sheet");
            var sheet2 = new Worksheet("sheet");

            var list = new WorksheetCollection();
            list.Add(sheet1);

            // Act
            try
            {
                list.Add(sheet2);
            }
            catch (ArgumentException) { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void WorksheetCollection_ByName_Get()
        {
            // Arrange
            var sheet = new Worksheet("sheet");

            var list = new WorksheetCollection();
            list.Add(sheet);

            // Act
            var returnedSheet = list["sheet"];

            // Assert
            Assert.AreEqual(sheet, returnedSheet);
        }

        [TestMethod]
        public void WorksheetCollection_ByName_Get_CacheWorks()
        {
            // Arrange
            var sheet = new Worksheet("sheet");

            var list = new WorksheetCollection();
            list.Add(sheet);
            var returnedSheet = list["sheet"];

            // Act
            var cachedSheet = list["sheet"];

            // Assert
            Assert.AreEqual(returnedSheet, cachedSheet);
        }

        [TestMethod]
        public void WorksheetCollection_ByName_Get_SheetNotFount_ReturnsNull()
        {
            // Arrange
            var sheet = new Worksheet("sheet");

            var list = new WorksheetCollection();
            list.Add(sheet);

            // Act
            var returnedSheet = list["wrongsheet"];

            // Assert
            Assert.IsNull(returnedSheet);
        }
    }
}
