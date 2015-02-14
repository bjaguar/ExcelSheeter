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
    public class NumberFormatStyleTest
    {
        [TestMethod]
        public void NumberFormatStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new NumberFormatStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void NumberFormatStyle_GetSet_Properties()
        {
            // Arrange
            var format = "format";
            var style = new NumberFormatStyle();

            // Act
            style.Format = format;

            // Assert
            Assert.AreEqual(format, style.Format);
        }

        [TestMethod]
        public void NumberFormatStyle_OuterXml()
        {
            // Arrange
            var format = "format";
            var style = new NumberFormatStyle();
            style.Format = format;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void NumberFormatStyle_OuterXml_EmptyNode()
        {
            // Arrange
            var style = new NumberFormatStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
            Assert.AreEqual(string.Empty, xml);
        }

        [TestMethod]
        public void NumerFormatStyle_ImplicitOperator()
        {
            // Arrange
            NumberFormatStyle style;

            // Act
            style = "format";
            string format = style;

            // Assert
            Assert.AreEqual("format", format);
        }

        [TestMethod]
        public void NumerFormatStyle_ImplicitOperator_EmptyStringParam_ReturnsNull()
        {
            // Arrange

            // Act
            NumberFormatStyle style = string.Empty;

            // Assert
            Assert.IsNull(style);
        }

        [TestMethod]
        public void NumerFormatStyle_ImplicitOperator_NullFormatParam()
        {
            // Arrange

            // Act
            NumberFormatStyle style = null;
            string format = style;

            // Assert
            Assert.AreEqual(string.Empty, format);
        }
    }
}
