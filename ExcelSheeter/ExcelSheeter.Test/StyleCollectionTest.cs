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
using System.Linq;
using System.Collections;

namespace ExcelSheeter.Test
{
    [TestClass]
    public class StyleCollectionTest
    {
        [TestMethod]
        public void StyleCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new StyleCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void StyleCollection_IsReadOnly()
        {
            // Arrange
            var list = new StyleCollection();

            // Act
            var value = list.IsReadOnly;

            // Assert
            Assert.IsFalse(value);
        }

        [TestMethod]
        public void StyleCollection_GetEnumerator()
        {
            // Arrange
            var list = new StyleCollection();

            // Act
            var enumerator = list.GetEnumerator();

            // Assert
            Assert.IsNotNull(enumerator);
        }

        [TestMethod]
        public void StyleCollection_GetEnumerator_2()
        {
            // Arrange
            var list = new StyleCollection();
            list.Add(new Style("id1"));
            list.Add(new Style("id2"));

            // Act
            var enumerator = (list as IEnumerable).GetEnumerator();

            // Assert
            Assert.IsNotNull(enumerator);
        }

        [TestMethod]
        public void StyleCollection_CopyTo()
        {
            // Arrange
            var style = new Style("id");

            var list = new StyleCollection();
            list.Add(style);
            
            var items = new Style[1];

            // Act
            list.CopyTo(items, 0);

            // Assert
            Assert.AreEqual(style, items[0]);
        }

        [TestMethod]
        public void StyleCollection_Add()
        {
            // Arrange
            var style = new Style("id");
            var list = new StyleCollection();

            // Act
            list.Add(style);

            // Assert
            Assert.AreEqual(1, list.Count);
        }

        [TestMethod]
        public void StyleCollection_Clear()
        {
            // Arrange
            var style = new Style("id");
            var list = new StyleCollection();
            list.Add(style);

            // Act
            list.Clear();

            // Assert
            Assert.AreEqual(0, list.Count);
        }

        [TestMethod]
        public void StyleCollection_Contains()
        {
            // Arrange
            var style = new Style("id");
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var value = list.Contains(style);

            // Assert
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void StyleCollection_Remove()
        {
            // Arrange
            var style = new Style("id");
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var value = list.Remove(style);

            // Assert
            Assert.AreEqual(0, list.Count);
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void StyleCollection_OuterXml()
        {
            // Arrange
            var style = new Style("id");
            var list = new StyleCollection();
            list.Add(style);

            // Act
            var xml = list.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void StyleCollection_OuterXml_EmptyXml()
        {
            // Arrange
            var list = new StyleCollection();

            // Act
            var xml = list.OuterXml;

            // Assert
            Assert.IsTrue(string.IsNullOrEmpty(xml));
        }
    }
}
