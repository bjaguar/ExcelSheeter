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
using System.Collections;

namespace ExcelSheeter.Test
{
    [TestClass]
    public class BorderStyleCollectionTest
    {
        [TestMethod]
        public void BorderStyleCollection_CreateNew()
        {
            // Arrange

            // Act
            var list = new BorderStyleCollection();

            // Assert
            Assert.IsNotNull(list);
        }

        [TestMethod]
        public void BorderStyleCollection_Add_ExistingBorder_OverritesItem()
        {
            // Arrange
            var list = new BorderStyleCollection();
            list.Add(BorderStylePosition.Bottom, "#000");

            // Act
            list.Add(BorderStylePosition.Bottom, "#fff");

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual("#fff", list[BorderStylePosition.Bottom].Color);
        }

        [TestMethod]
        public void BorderStyleCollection_Add_WithParams_Position()
        {
            // Arrange
            var position = BorderStylePosition.Left;

            var list = new BorderStyleCollection();

            // Act
            list.Add(position, "#000");

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(position, list[BorderStylePosition.Left].Position);
        }

        [TestMethod]
        public void BorderStyleCollection_Add_WithParams_Position_Color()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";

            var list = new BorderStyleCollection();

            // Act
            list.Add(position, color);

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(position, list[BorderStylePosition.Left].Position);
            Assert.AreEqual(color, list[BorderStylePosition.Left].Color);
        }

        [TestMethod]
        public void BorderStyleCollection_Add_WithParams_Position_Color_LineStyle()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";
            var lineStyle = BorderLineStyle.SlantDashDot;

            var list = new BorderStyleCollection();

            // Act
            list.Add(position, color, lineStyle);

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(position, list[BorderStylePosition.Left].Position);
            Assert.AreEqual(color, list[BorderStylePosition.Left].Color);
            Assert.AreEqual(lineStyle, list[BorderStylePosition.Left].LineStyle);
        }

        [TestMethod]
        public void BorderStyleCollection_Add_WithParams_Position_Color_LineStyle_Weight()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";
            var lineStyle = BorderLineStyle.SlantDashDot;
            var weight = 10;

            var list = new BorderStyleCollection();

            // Act
            list.Add(position, color, lineStyle, weight);

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(position, list[BorderStylePosition.Left].Position);
            Assert.AreEqual(color, list[BorderStylePosition.Left].Color);
            Assert.AreEqual(lineStyle, list[BorderStylePosition.Left].LineStyle);
            Assert.AreEqual(weight, list[BorderStylePosition.Left].Weight);
        }

        [TestMethod]
        public void BorderStyleCollection_GetEnumerator()
        {
            // Arrange
            var list = new BorderStyleCollection();

            // Act
            var enumerator = list.GetEnumerator();

            // Assert
            Assert.IsNotNull(enumerator);
        }

        [TestMethod]
        public void BorderStyleCollection_GetEnumerator_2()
        {
            // Arrange
            var list = new BorderStyleCollection();

            // Act
            var enumerator = (list as IEnumerable).GetEnumerator();

            // Assert
            Assert.IsNotNull(enumerator);
        }

        [TestMethod]
        public void BorderStyleCollection_Get_BorderNotFound_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var list = new BorderStyleCollection();

            // Act
            try
            {
                var border = list[BorderStylePosition.Left];
            }
            catch { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void BorderStyleCollection_Remove()
        {
            // Arrange
            var list = new BorderStyleCollection();
            list.Add(BorderStylePosition.Bottom, "#000");

            // Act
            var value = list.Remove(BorderStylePosition.Bottom);

            // Assert
            Assert.IsTrue(value);
        }

        [TestMethod]
        public void BorderStyleCollection_Remove_ItemNotFound_ReturnsFalse()
        {
            // Arrange
            var list = new BorderStyleCollection();
            list.Add(BorderStylePosition.Bottom, "#000");

            // Act
            var value = list.Remove(BorderStylePosition.Top);

            // Assert
            Assert.IsFalse(value);
        }
    }
}
