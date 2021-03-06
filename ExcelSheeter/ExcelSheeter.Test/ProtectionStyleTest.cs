﻿/*
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
    public class ProtectionStyleTest
    {
        [TestMethod]
        public void ProtectionStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new ProtectionStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void ProtectionStyle_GetSet_Properties_TrueValues()
        {
            // Arrange
            var @protected = true;
            var hideFormula = true;

            var style = new ProtectionStyle();

            // Act
            style.Protected = @protected;
            style.HideFormula = hideFormula;

            // Assert
            Assert.AreEqual(@protected, style.Protected);
            Assert.AreEqual(hideFormula, style.HideFormula);
        }

        [TestMethod]
        public void ProtectionStyle_GetSet_Properties_FalseValues()
        {
            // Arrange
            var @protected = false;
            var hideFormula = false;

            var style = new ProtectionStyle();

            // Act
            style.Protected = @protected;
            style.HideFormula = hideFormula;

            // Assert
            Assert.AreEqual(@protected, style.Protected);
            Assert.AreEqual(hideFormula, style.HideFormula);
        }

        [TestMethod]
        public void ProtectionStyle_OuterXml()
        {
            // Arrange
            var @protected = true;
            var hideFormula = false;

            var style = new ProtectionStyle();
            style.Protected = @protected;
            style.HideFormula = hideFormula;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void ProtectionStyle_OuterXml_EmptyXml()
        {
            // Arrange
            var style = new ProtectionStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsTrue(string.IsNullOrEmpty(xml));
        }
    }
}
