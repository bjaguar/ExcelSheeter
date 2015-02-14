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
    public class AlignmentStyleTest
    {
        [TestMethod]
        public void AlignmentStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new AlignmentStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void AlignmentStyle_GetSet_Properties()
        {
            // Arrange
            var horizontal = HorizontalAlignment.Center;
            var indent = 10;
            var readingOrder = ReadingOrder.LeftToRight;
            var rotate = 0;
            var shrinkToFit = false;
            var vertical = VerticalAlignment.Center;
            var verticalText = false;
            var wrapText = true;

            var style = new AlignmentStyle();

            // Act
            style.Horizontal = horizontal;
            style.Indent = indent;
            style.ReadingOrder = readingOrder;
            style.Rotate = rotate;
            style.ShrinkToFit = shrinkToFit;
            style.Vertical = vertical;
            style.VerticalText = verticalText;
            style.WrapText = wrapText;

            // Assert
            Assert.AreEqual(horizontal, style.Horizontal); 
            Assert.AreEqual(indent, style.Indent); 
            Assert.AreEqual(readingOrder, style.ReadingOrder); 
            Assert.AreEqual(rotate, style.Rotate); 
            Assert.AreEqual(shrinkToFit, style.ShrinkToFit); 
            Assert.AreEqual(vertical, style.Vertical); 
            Assert.AreEqual(verticalText, style.VerticalText); 
            Assert.AreEqual(wrapText, style.WrapText); 
        }

        [TestMethod]
        public void AlignmentStyle_ShrinkToFitProperty_TrueValue()
        {
            // Arrange
            var shrinkToFit = true;

            var style = new AlignmentStyle();

            // Act
            style.ShrinkToFit = shrinkToFit;

            // Assert
            Assert.AreEqual(shrinkToFit, style.ShrinkToFit);
        }

        [TestMethod]
        public void AlignmentStyle_VerticalTextProperty_TrueValue()
        {
            // Arrange
            var verticalText = true;

            var style = new AlignmentStyle();

            // Act
            style.VerticalText = verticalText;

            // Assert
            Assert.AreEqual(verticalText, style.VerticalText);
        }

        [TestMethod]
        public void AlignmentStyle_WrapTextProperty_FalseValue()
        {
            // Arrange
            var wrapText = false;

            var style = new AlignmentStyle();

            // Act
            style.WrapText = wrapText;

            // Assert
            Assert.AreEqual(wrapText, style.WrapText);
        }

        [TestMethod]
        public void AlignmentStyle_OuterXml()
        {
            // Arrange
            var horizontal = HorizontalAlignment.Center;
            var indent = 10;
            var readingOrder = ReadingOrder.LeftToRight;
            var rotate = 0;
            var shrinkToFit = false;
            var vertical = VerticalAlignment.Center;
            var verticalText = false;
            var wrapText = true;

            var style = new AlignmentStyle();
            style.Horizontal = horizontal;
            style.Indent = indent;
            style.ReadingOrder = readingOrder;
            style.Rotate = rotate;
            style.ShrinkToFit = shrinkToFit;
            style.Vertical = vertical;
            style.VerticalText = verticalText;
            style.WrapText = wrapText;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }

        [TestMethod]
        public void AlignmentStyle_OuterXml_EmptyXml()
        {
            // Arrange
            var style = new AlignmentStyle();

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsTrue(string.IsNullOrWhiteSpace(xml));
        }
    }
}
