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
        public void BorderStyleCollection_Add_WithParams_Position()
        {
            // Arrange
            var position = BorderStylePosition.Left;

            var list = new BorderStyleCollection();

            // Act
            list.Add(position);

            // Assert
            Assert.AreEqual(1, list.Count);
            Assert.AreEqual(position, list[0].Position);
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
            Assert.AreEqual(position, list[0].Position);
            Assert.AreEqual(color, list[0].Color);
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
            Assert.AreEqual(position, list[0].Position);
            Assert.AreEqual(color, list[0].Color);
            Assert.AreEqual(lineStyle, list[0].LineStyle);
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
            Assert.AreEqual(position, list[0].Position);
            Assert.AreEqual(color, list[0].Color);
            Assert.AreEqual(lineStyle, list[0].LineStyle);
            Assert.AreEqual(weight, list[0].Weight);
        }
    }
}
