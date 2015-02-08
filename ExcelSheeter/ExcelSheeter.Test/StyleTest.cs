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
    public class StyleTest
    {
        [TestMethod]
        public void Style_CreateNew()
        {
            // Arrange

            // Act
            var style = new Style();

            // Assert
            Assert.IsNotNull(style);
            Assert.IsNotNull(style.AlignmentStyle);
            Assert.IsNotNull(style.BordersStyle);
            Assert.IsNotNull(style.FontStyle);
            Assert.IsNotNull(style.InteriorStyle);
            Assert.IsNotNull(style.NumberFormatStyle);
            Assert.IsNotNull(style.ProtectionStyle);
        }

        [TestMethod]
        public void Style_GetSet_Properties()
        {
            // Arrange
            var id = "id";
            var name = "name";

            var style = new Style();

            // Act
            style.Id = id;
            style.Name = name;

            // Assert
            Assert.AreEqual(id, style.Id);
            Assert.AreEqual(name, style.Name);
        }

        [TestMethod]
        public void Style_Set_EmptyId_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var id = "";

            var style = new Style();

            // Act
            try
            {
                style.Id = id;
            }
            catch (ArgumentNullException) { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void Style_Set_EmptyName_ThrowsException()
        {
            // Arrange
            var thrown = false;
            var name = "";

            var style = new Style();

            // Act
            try
            {
                style.Name = name;
            }
            catch (ArgumentNullException) { thrown = true; }

            // Assert
            Assert.IsTrue(thrown);
        }

        [TestMethod]
        public void Style_OuterXml()
        {
            // Arrange
            var id = "id";
            var name = "name";

            var style = new Style();
            style.Id = id;
            style.Name = name;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
