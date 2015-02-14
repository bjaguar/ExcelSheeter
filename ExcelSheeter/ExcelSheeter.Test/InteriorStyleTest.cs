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
    public class InteriorStyleTest
    {
        [TestMethod]
        public void InteriorStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new InteriorStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void InteriorStyle_GetSet_Properties()
        {
            // Arrange
            var color = "#fff";
            var pattern = InteriorStylePattern.Gray0625;
            var patternColor = "#333";

            var style = new InteriorStyle();

            // Act
            style.Color = color;
            style.Pattern = pattern;
            style.PatternColor = patternColor;

            // Assert
            Assert.AreEqual(color, style.Color);
            Assert.AreEqual(pattern, style.Pattern);
            Assert.AreEqual(patternColor, style.PatternColor);
        }

        [TestMethod]
        public void InteriorStyle_ColorProperty_EmptyValue_RemovesAttribute()
        {
            // Arrange
            var color = "#fff";

            var style = new InteriorStyle();
            style.Color = color;

            // Act
            style.Color = string.Empty;

            // Assert
            Assert.AreEqual(0, style.Attributes.Count);
        }

        [TestMethod]
        public void InteriorStyle_PatternColorProperty_EmptyValue_RemovesAttribute()
        {
            // Arrange
            var patternColor = "#fff";

            var style = new InteriorStyle();
            style.PatternColor = patternColor;

            // Act
            style.PatternColor = string.Empty;

            // Assert
            Assert.AreEqual(0, style.Attributes.Count);
        }

        [TestMethod]
        public void InteriorStyle_OuterXml()
        {
            // Arrange
            var color = "#fff";
            var pattern = InteriorStylePattern.Gray0625;
            var patternColor = "#333";

            var style = new InteriorStyle();
            style.Color = color;
            style.Pattern = pattern;
            style.PatternColor = patternColor;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void InteriorStyle_OuterXml_EmptyXml()
        {
            // Arrange
            var style = new InteriorStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsTrue(string.IsNullOrEmpty(xml));
        }
    }
}
