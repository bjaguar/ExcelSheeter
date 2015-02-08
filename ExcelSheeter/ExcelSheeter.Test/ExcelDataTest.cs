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
    public class ExcelDataTest
    {
        [TestMethod]
        public void ExcelData_CreateNew()
        {
            // Arrange

            // Act
            var data = new ExcelData();

            // Assert
            Assert.IsNotNull(data);
        }

        [TestMethod]
        public void ExcelData_CreateNew_WithParams_Value()
        {
            // Arrange
            var value = "value";

            // Act
            var data = new ExcelData(value);

            // Assert
            Assert.IsNotNull(data);
            Assert.AreEqual(value, data.Value);
        }

        [TestMethod]
        public void ExcelData_CreateNew_WithParams_ValueAndDataType()
        {
            // Arrange
            var value = 1;
            var dataType = DataType.Number;

            // Act
            var data = new ExcelData(dataType, value);

            // Assert
            Assert.IsNotNull(data);
            Assert.AreEqual(dataType, data.DataType);
            Assert.AreEqual(value, data.Value);
        }

        [TestMethod]
        public void ExcelData_OuterXml()
        {
            // Arrange
            var value = "value";
            var data = new ExcelData(value);

            // Act
            var xml = data.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
