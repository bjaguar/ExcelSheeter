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
    public class FontStyleTest
    {
        [TestMethod]
        public void FontStyle_CreateNew()
        {
            // Arrange

            // Act
            var style = new FontStyle();

            // Assert
            Assert.IsNotNull(style);
        }

        [TestMethod]
        public void FontStyle_GetSet_Properties()
        {
            // Arrange
            var bold = false;
            var color = "#fff";
            var fontName = "Arial";
            var italic = false;
            var outline = true;
            var shadow = false;
            var size = 10;
            var strikeThrough = false;
            var underline = FontUnderline.Single;
            var verticalAlign = FontVerticalAlign.None;
            var charSet = 12;
            var family = FontFamily.Modern;

            var style = new FontStyle();

            // Act
            style.Bold = bold;
            style.Color = color;
            style.FontName = fontName;
            style.Italic = italic;
            style.Outline = outline;
            style.Shadow = shadow;
            style.Size = size;
            style.StrikeThrough = strikeThrough;
            style.Underline = underline;
            style.VerticalAlign = verticalAlign;
            style.CharSet = charSet;
            style.Family = family;

            // Assert
            Assert.AreEqual(bold, style.Bold);
            Assert.AreEqual(color, style.Color);
            Assert.AreEqual(fontName, style.FontName);
            Assert.AreEqual(italic, style.Italic);
            Assert.AreEqual(outline, style.Outline);
            Assert.AreEqual(shadow, style.Shadow);
            Assert.AreEqual(size, style.Size);
            Assert.AreEqual(strikeThrough, style.StrikeThrough);
            Assert.AreEqual(underline, style.Underline);
            Assert.AreEqual(verticalAlign, style.VerticalAlign);
            Assert.AreEqual(charSet, style.CharSet);
            Assert.AreEqual(family, style.Family);
        }

        [TestMethod]
        public void FontStyle_OuterXml()
        {
            // Arrange
            var bold = false;
            var color = "#fff";
            var fontName = "Arial";
            var italic = false;
            var outline = true;
            var shadow = false;
            var size = 10;
            var strikeThrough = false;
            var underline = FontUnderline.Single;
            var verticalAlign = FontVerticalAlign.None;
            var charSet = 12;
            var family = FontFamily.Modern;

            var style = new FontStyle();
            style.Bold = bold;
            style.Color = color;
            style.FontName = fontName;
            style.Italic = italic;
            style.Outline = outline;
            style.Shadow = shadow;
            style.Size = size;
            style.StrikeThrough = strikeThrough;
            style.Underline = underline;
            style.VerticalAlign = verticalAlign;
            style.CharSet = charSet;
            style.Family = family;

            // Act
            var xml = style.OuterXml;

            // Assert
            Assert.IsNotNull(xml);
        }
    }
}
