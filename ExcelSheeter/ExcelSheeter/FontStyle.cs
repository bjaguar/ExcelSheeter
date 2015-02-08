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
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents the cell's font.
    /// </summary>
    public sealed class FontStyle : StyleEntity
    {
        /// <summary>
        /// Gets or sets a value indicating if the font is bold.
        /// </summary>
        public bool Bold
        {
            get
            {
                var value = Attributes[AttributeConstants.BoldName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.BoldName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets the font's color.
        /// </summary>
        public string Color
        {
            get { return Attributes[AttributeConstants.ColorName]; }
            set { Attributes[AttributeConstants.ColorName] = value; }
        }

        /// <summary>
        /// Gets or sets the font's name.
        /// </summary>
        public string FontName
        {
            get { return Attributes[AttributeConstants.FontName]; }
            set { Attributes[AttributeConstants.FontName] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating if the font is italic.
        /// </summary>
        public bool Italic
        {
            get
            {
                var value = Attributes[AttributeConstants.ItalicName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.ItalicName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value indicating if the font is outline.
        /// </summary>
        public bool Outline
        {
            get
            {
                var value = Attributes[AttributeConstants.OutlineName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.OutlineName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value indicating if the font is shadowed.
        /// </summary>
        public bool Shadow
        {
            get
            {
                var value = Attributes[AttributeConstants.ShadowName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.ShadowName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets the font's size.
        /// </summary>
        public double Size
        {
            get { return double.Parse(Attributes[AttributeConstants.SizeName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.SizeName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value indicating if the font is strike through.
        /// </summary>
        public bool StrikeThrough
        {
            get
            {
                var value = Attributes[AttributeConstants.StrikeThroughName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.StrikeThroughName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets the font's underline type.
        /// </summary>
        public FontUnderline Underline
        {
            get { return (FontUnderline)Enum.Parse(typeof(FontUnderline), Attributes[AttributeConstants.UnderlineName]); }
            set { Attributes[AttributeConstants.UnderlineName] = Enum.GetName(typeof(FontUnderline), value); }
        }

        /// <summary>
        /// Gets or sets the font's vertical align.
        /// </summary>
        public FontVerticalAlign VerticalAlign
        {
            get { return (FontVerticalAlign)Enum.Parse(typeof(FontVerticalAlign), Attributes[AttributeConstants.VerticalAlignName]); }
            set { Attributes[AttributeConstants.VerticalAlignName] = Enum.GetName(typeof(FontVerticalAlign), value); }
        }

        /// <summary>
        /// Gets or sets the font's charset.
        /// </summary>
        public long CharSet
        {
            get { return long.Parse(Attributes[AttributeConstants.CharSetName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.CharSetName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the font family.
        /// </summary>
        public FontFamily Family
        {
            get { return (FontFamily)Enum.Parse(typeof(FontFamily), Attributes[AttributeConstants.FamilyName]); }
            set { Attributes[AttributeConstants.FamilyName] = Enum.GetName(typeof(FontFamily), value); }
        }

        internal override string TagName { get { return "ss:Font"; } }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml
        {
            get
            {
                if (Attributes.Count > 0) return base.OuterXml;

                return string.Empty;
            }
        }
    }
}
