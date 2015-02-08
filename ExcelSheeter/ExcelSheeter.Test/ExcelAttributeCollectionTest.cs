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
    public class ExcelAttributeCollectionTest
    {
        [TestMethod]
        public void ExcelAttributeCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new ExcelAttributeCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void ExcelAttributeCollection_Add()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();

            // Act
            list.Add(key, value);

            // Assert
            Assert.AreEqual(1, list.Count);
        }

        [TestMethod]
        public void ExcelAttributeCollection_Contains_ExistingKey()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            var exists = list.Contains("key");

            // Assert
            Assert.IsTrue(exists);
        }

        [TestMethod]
        public void ExcelAttributeCollection_Contains_NotExistingKey()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            var exists = list.Contains("wrongkey");

            // Assert
            Assert.IsFalse(exists);
        }

        [TestMethod]
        public void ExcelAttributeCollection_Remove_ByKey_ReturnsTrueIfItemExists()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            var removed = list.Remove("key");

            // Assert
            Assert.AreEqual(0, list.Count);
            Assert.IsTrue(removed);
        }

        [TestMethod]
        public void ExcelAttributeCollection_Remove_ByKey_ReturnsFalseIfItemExists()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            var removed = list.Remove("wrongkey");

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.IsFalse(removed);
        }

        [TestMethod]
        public void ExcelAttributeCollection_ByKey_GetsValue()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            var itemValue = list["key"];

            // Assert
            Assert.AreEqual(value, itemValue);
        }

        [TestMethod]
        public void ExcelAttributeCollection_ByKey_SetsValue_IfExists()
        {
            // Arrange
            var key = "key";
            var value = "value";

            var list = new ExcelAttributeCollection();

            // Act
            list[key] = value;

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(value, list[0].Value);
        }

        [TestMethod]
        public void ExcelAttributeCollection_ByKey_SetsValue_CreateIfNotExists()
        {
            // Arrange
            var key = "key";
            var value = "value";
            var modifiedValue = "new value";

            var list = new ExcelAttributeCollection();
            list.Add(key, value);

            // Act
            list[key] = modifiedValue;

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(modifiedValue, list[0].Value);
        }
    }
}
