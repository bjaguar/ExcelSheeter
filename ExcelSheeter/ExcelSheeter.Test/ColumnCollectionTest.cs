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
    public class ColumnCollectionTest
    {
        [TestMethod]
        public void ColumnCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new ColumnCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void ColumnCollection_SetByIndex()
        {
            // Arrange
            var collection = new Column();
            var list = new ColumnCollection();

            // Act
            list[1] = collection;

            // Assert
            Assert.AreEqual(2, list.Count);
        }

        [TestMethod]
        public void ColumnCollection_GetByIndex()
        {
            // Arrange
            var collection = new Column();
            var list = new ColumnCollection();
            list[1] = collection;

            // Act
            var returnedCollection = list[1];

            // Assert
            Assert.AreEqual(collection, returnedCollection);
        }
    }
}
