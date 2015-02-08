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
    /// Represents an Excel row.
    /// </summary>
    public sealed class Row : ExcelEntity
    {
        private readonly CellCollection cells = new CellCollection();

        /// <summary>
        /// Creates a new instance of the <see cref="Row"/> object.
        /// </summary>
        public Row()
        {
        }

        /// <summary>
        /// Returns a cell given the column name.
        /// </summary>
        /// <param name="columnName">Column name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName)) throw new ArgumentNullException("columnName");

            int index = -1;

            var columnNameUpperName = columnName.ToUpper(CultureInfo.InvariantCulture);

            for (int i = 0; i < columnNameUpperName.Length; i++) index += (Char.ConvertToUtf32(columnNameUpperName, i) - 64) + (25 * i);

            var cell = GetCell(index);

            return cell;
        }

        /// <summary>
        /// Returns a cell given it's index.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(int columnIndex)
        {
            var cell = Cells[columnIndex];

            return cell;
        }

        /// <summary>
        /// Gets or sets the value of the caption.
        /// </summary>
        public string Caption
        {
            get { return Attributes[AttributeConstants.CaptionName]; }
            set { Attributes[AttributeConstants.CaptionName] = value; }
        }

        /// <summary>
        /// Gets or sets a value that indicates if the row be autosized for numeric and date values only.
        /// </summary>
        /// <remarks>
        /// From Microsoft Reference:
        /// <list type="bullet">
        /// <item>ss:AutoFitWidth="1" and ss:Width is unspecified: Autofit the column width to fit the content.</item>
        /// <item>ss:AutoFitWidth="1" and ss:Width is specified: Set the column to the specified width and only autofit if the size of the content is larger than the specified width.</item>
        /// <item>ss:AutoFitWidth="0" and ss:Width is unspecified: Use the default column width.</item>
        /// <item>ss:AutoFitWidth="0" and ss:Width is specified: Use the specified width.</item>
        /// </list>
        /// </remarks>
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
        /// Gets or sets a value that specifies the height of a row in points.
        /// </summary>
        public double Height
        {
            get { return double.Parse(Attributes[AttributeConstants.HeightName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.HeightName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value that indicates if the column will be displayed.
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
        /// Gets or sets the index of the row.
        /// </summary>
        public long Index
        {
            get { return long.Parse(Attributes[AttributeConstants.IndexName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.IndexName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value that specifies the number of adjacent rows with the same formatting as this row.
        /// </summary>
        public long Span
        {
            get { return long.Parse(Attributes[AttributeConstants.SpanName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.SpanName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the identifier of the style applied to this row.
        /// </summary>
        public string StyleId
        {
            get { return Attributes[AttributeConstants.StyleIdName]; }
            set { Attributes[AttributeConstants.StyleIdName] = value; }
        }

        /// <summary>
        /// Gets the cell collection.
        /// </summary>
        public CellCollection Cells { get { return cells; } }

        internal override string TagName { get { return "ss:Row"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                foreach (var cell in Cells) xml += cell.OuterXml;

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml { get { return base.OuterXml + "\r\n"; } }
    }
}
