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
    public class BordersStyleTest
    {
        [TestMethod]
        public void BordersStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new BordersStyle();

            // Assert
            Assert.IsNotNull(style);
            Assert.IsNotNull(style.Borders);
        }

        [TestMethod]
        public void BordersStyle_OuterXml()
        {
            // Arrange
            var style = new BordersStyle();
            style.Borders.Add(BorderStylePosition.Bottom, "#000");

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void BordersStyle_OuterXml_EmptyXml()
        {
            // Arrange
            var style = new BordersStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsTrue(string.IsNullOrWhiteSpace(xml));
        }
    }
}
