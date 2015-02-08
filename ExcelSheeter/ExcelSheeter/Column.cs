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
    /// Represents an Excel column.
    /// </summary>
    public sealed class Column : ExcelEntity
    {
        /// <summary>
        /// Creates a new <see cref="Column"/> object.
        /// </summary>
        public Column()
        {
        }

        /// <summary>
        /// Gets or sets the column's caption.
        /// </summary>
        public string Caption
        {
            get { return Attributes[AttributeConstants.CaptionName]; }
            set { Attributes[AttributeConstants.CaptionName] = value; }
        }

        /// <summary>
        /// Gets or sets a value that indicates if the column adjust it's with to the content.
        /// </summary>
        public bool AutoFitWidth
        {
            get
            {
                var value = Attributes[AttributeConstants.AutoFitWidthName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.AutoFitWidthName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value indicating if the column is visible.
        /// </summary>
        public bool Hidden
        {
            get
            {
                var value = Attributes[AttributeConstants.HiddenName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.HiddenName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets the column's index.
        /// </summary>
        public long Index
        {
            get { return long.Parse(Attributes[AttributeConstants.IndexName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.IndexName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value that specifies the number of adjacent columns with the same formatting as this column.
        /// </summary>
        public long Span
        {
            get { return long.Parse(Attributes[AttributeConstants.SpanName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.SpanName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the identifier of the column style.
        /// </summary>
        public string StyleId
        {
            get { return Attributes[AttributeConstants.StyleIdName]; }
            set { Attributes[AttributeConstants.StyleIdName] = value; }
        }

        /// <summary>
        /// Gets or sets the column's width.
        /// </summary>
        public double Width
        {
            get { return double.Parse(Attributes[AttributeConstants.WidthName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.WidthName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        internal override string TagName { get { return "ss:Column"; } }
    }
}
