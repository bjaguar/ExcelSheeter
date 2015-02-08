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
    public class ColumnTest
    {
        [TestMethod]
        public void Column_CreateNew()
        {
            // Arrange

            // Act
            var column = new Column();

            //Assert
            Assert.IsNotNull(column);
        }

        [TestMethod]
        public void Column_GetAndSet_Properties()
        {
            // Arrange
            var caption = "caption";
            var autoFitWidth = true;
            var hidden = false;
            var index = 2;
            var span = 1;
            var styleId = "id";
            var width = 100;

            var column = new Column();

            // Act
            column.Caption = caption;
            column.AutoFitWidth = autoFitWidth;
            column.Hidden = hidden;
            column.Index = index;
            column.Span = span;
            column.StyleId = styleId;
            column.Width = width;

            //Assert
            Assert.AreEqual(caption, column.Caption);
            Assert.AreEqual(autoFitWidth, column.AutoFitWidth);
            Assert.AreEqual(hidden, column.Hidden);
            Assert.AreEqual(index, column.Index);
            Assert.AreEqual(span, column.Span);
            Assert.AreEqual(styleId, column.StyleId);
            Assert.AreEqual(width, column.Width);
        }
    }
}
