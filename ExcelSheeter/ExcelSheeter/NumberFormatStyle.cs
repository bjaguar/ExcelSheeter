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

namespace ExcelSheeter
{
    /// <summary>
    /// Represents the number format style.
    /// </summary>
    public sealed class NumberFormatStyle : StyleEntity
    {
        /// <summary>
        /// Converts an <see cref="NumberFormatStyle"/> to string.
        /// </summary>
        /// <param name="style">A <see cref="NumberFormatStyle"/> object.</param>
        /// <returns>A string with the format.</returns>
        public static implicit operator string(NumberFormatStyle style)
        {
            if (style != null) return style.Format;

            return string.Empty;
        }

        /// <summary>
        /// Converts a string into a <see cref="NumberFormatStyle"/> object.
        /// </summary>
        /// <param name="value">The format's value.</param>
        /// <returns>A <see cref="NumberFormatStyle"/> object.</returns>
        public static implicit operator NumberFormatStyle(string value)
        {
            if (!string.IsNullOrWhiteSpace(value)) return new NumberFormatStyle { Format = value, };

            return null;
        }

        /// <summary>
        /// Gets or sets the format style value.
        /// </summary>
        internal string Format
        {
            get { return Attributes[AttributeConstants.FormatName]; }
            set { Attributes[AttributeConstants.FormatName] = value; }
        }

        internal override string TagName { get { return "ss:NumberFormat"; } }

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
