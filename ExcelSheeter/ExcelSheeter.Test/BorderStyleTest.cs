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
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelSheeter.Test
{
    [TestClass]
    public class BorderStyleTest
    {
        [TestMethod]
        public void BorderStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new BorderStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void BorderStyle_CreateNew_WithParams_Position()
        {
            // Arrange
            var position = BorderStylePosition.Left;

            // Act
            var style = new BorderStyle(position);

            // Assert
            Assert.IsNotNull(style);
            Assert.AreEqual(position, style.Position);
        }

        [TestMethod]
        public void BorderStyle_CreateNew_WithParams_Position_Color()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";

            // Act
            var style = new BorderStyle(position, color);

            // Assert
            Assert.IsNotNull(style);
            Assert.AreEqual(position, style.Position);
            Assert.AreEqual(color, style.Color);
        }

        [TestMethod]
        public void BorderStyle_CreateNew_WithParams_Position_Color_BorderLineStyle()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";
            var borderLineStyle = BorderLineStyle.Double;

            // Act
            var style = new BorderStyle(position, color, borderLineStyle);

            // Assert
            Assert.IsNotNull(style);
            Assert.AreEqual(position, style.Position);
            Assert.AreEqual(color, style.Color);
            Assert.AreEqual(borderLineStyle, style.LineStyle);
        }

        [TestMethod]
        public void BorderStyle_CreateNew_WithParams_Position_Color_BorderLineStyle_Weight()
        {
            // Arrange
            var position = BorderStylePosition.Left;
            var color = "#fff";
            var borderLineStyle = BorderLineStyle.Double;
            var weight = 10;

            // Act
            var style = new BorderStyle(position, color, borderLineStyle, weight);

            // Assert
            Assert.IsNotNull(style);
            Assert.AreEqual(position, style.Position);
            Assert.AreEqual(color, style.Color);
            Assert.AreEqual(borderLineStyle, style.LineStyle);
            Assert.AreEqual(weight, style.Weight);
        }

        [TestMethod]
        public void BorderStyle_OuterXml()
        {
            // Arrange
            var style = new BorderStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void BorderStyle_ColorProperty_EmptyValue_RemovesAttribute()
        {
            // Arrange
            var style = new BorderStyle();
            style.Color = "#fff";

            // Act
            style.Color = string.Empty;

            // Assert
            Assert.IsFalse(style.Attributes.Any(x => x.Key == AttributeConstants.ColorName));
        }
    }
}
