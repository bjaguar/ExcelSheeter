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
    public class ExcelAttributeTest
    {
        [TestMethod]
        public void ExcelAttribute_CreateNew()
        {
            // Arrange

            // Act
            var attr = new ExcelAttribute();

            // Assert
            Assert.IsNotNull(attr);
        }

        [TestMethod]
        public void ExcelAttribute_CreateNew_WithParams()
        {
            // Arrange
            var key = "attrKey";
            var value = "attribute value";

            // Act
            var attr = new ExcelAttribute(key, value);

            // Assert
            Assert.IsNotNull(attr);
            Assert.AreEqual(key, attr.Key);
            Assert.AreEqual(value, attr.Value);
        }

        [TestMethod]
        public void ExcelAttribute_SetValue()
        {
            // Arrange
            var attr = new ExcelAttribute("name", "value");

            // Act
            attr.Value = "newvalue";

            // Assert
            Assert.AreEqual("newvalue", attr.Value);
        }

        [TestMethod]
        public void ExcelAttribute_GetValue()
        {
            // Arrange
            var attr = new ExcelAttribute("name", "value");

            // Act
            var value = attr.Value;

            // Assert
            Assert.AreEqual("value", value);
        }
    }
}
