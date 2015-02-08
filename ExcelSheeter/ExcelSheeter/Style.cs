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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents an Excel style.
    /// </summary>
    public sealed class Style : ExcelEntity
    {
        private AlignmentStyle excelAlignmentStyle = new AlignmentStyle();

        private BordersStyle excelBordersStyle = new BordersStyle();

        private FontStyle excelFontStyle = new FontStyle();

        private InteriorStyle excelInteriorStyle = new InteriorStyle();

        private NumberFormatStyle excelNumberFormatStyle = new NumberFormatStyle();

        private ProtectionStyle excelProtectionStyle = new ProtectionStyle();

        /// <summary>
        /// Creates a new <see cref="Style"/> object.
        /// </summary>
        public Style()
        {
        }

        /// <summary>
        /// Gets or sets the style's identifier.
        /// </summary>
        public string Id
        {
            get { return Attributes[AttributeConstants.IdName]; }
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentNullException("value");

                Attributes[AttributeConstants.IdName] = value;
            }
        }

        /// <summary>
        /// Gets or sets the style's name.
        /// </summary>
        public string Name
        {
            get { return Attributes[AttributeConstants.Name]; }
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentNullException("value");

                Attributes[AttributeConstants.Name] = value;
            }
        }

        /// <summary>
        /// Gets the alignment style.
        /// </summary>
        public AlignmentStyle AlignmentStyle
        {
            get { return excelAlignmentStyle; }
        }

        /// <summary>
        /// Gets the borders style.
        /// </summary>
        public BordersStyle BordersStyle
        {
            get { return excelBordersStyle; }
        }

        /// <summary>
        /// Gets the font style.
        /// </summary>
        public FontStyle FontStyle
        {
            get { return excelFontStyle; }
        }

        /// <summary>
        /// Gets the interior style.
        /// </summary>
        public InteriorStyle InteriorStyle
        {
            get { return excelInteriorStyle; }
        }

        /// <summary>
        /// Gets the number format style.
        /// </summary>
        public NumberFormatStyle NumberFormatStyle
        {
            get { return excelNumberFormatStyle; }
        }

        /// <summary>
        /// Gets the protection style.
        /// </summary>
        public ProtectionStyle ProtectionStyle
        {
            get { return excelProtectionStyle; }
        }

        internal override string TagName { get { return "ss:Style"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                xml += excelAlignmentStyle.OuterXml;
                xml += excelBordersStyle.OuterXml;
                xml += excelFontStyle.OuterXml;
                xml += excelInteriorStyle.OuterXml;
                xml += excelNumberFormatStyle.OuterXml;
                xml += excelProtectionStyle.OuterXml;

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml { get { return base.OuterXml + "\r\n"; } }
    }
}
