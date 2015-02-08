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
using System.Collections.Generic;

namespace ExcelSheeter.Test
{
    [TestClass]
    public class ExcelEntityTest
    {
        private class Entity01 : ExcelEntity { internal override string TagName { get { return "xx:tag"; } } }

        private class Entity02 : ExcelEntity
        {
            public List<ExcelEntity> Items { get; set; }

            internal override string TagName { get { return "xx:entity"; } }

            public override string InnerXml
            {
                get
                {
                    var xml = string.Empty;

                    foreach (var item in Items) xml += item.OuterXml;

                    return xml;
                }
            }
        }

        [TestMethod]
        public void ExcelEntity_CheckOuterXml_01()
        {
            // Arrange
            var expectedXml = @"<xx:tag />";

            var entity = new Entity01();

            // Act
            var outerXml = entity.OuterXml;

            // Assert
            Assert.AreEqual(expectedXml, outerXml);
        }

        [TestMethod]
        public void ExcelEntity_CheckOuterXml_02()
        {
            // Arrange
            var expectedXml = @"<xx:tag id=""id"" name=""name"" />";

            var entity = new Entity01();

            entity.Attributes.Add("id", "id");
            entity.Attributes.Add("name", "name");

            // Act
            var outerXml = entity.OuterXml;

            // Assert
            Assert.AreEqual(expectedXml, outerXml);
        }

        [TestMethod]
        public void ExcelEntity_CheckOuterXml_03()
        {
            // Arrange
            var expectedXml = @"<xx:entity><xx:tag id=""id"" name=""name"" /></xx:entity>";

            var entity01 = new Entity01();

            entity01.Attributes.Add("id", "id");
            entity01.Attributes.Add("name", "name");

            var entity02 = new Entity02();
            entity02.Items = new List<ExcelEntity>() { entity01 };

            // Act
            var outerXml = entity02.OuterXml;

            // Assert
            Assert.AreEqual(expectedXml, outerXml);
        }
    }
}
