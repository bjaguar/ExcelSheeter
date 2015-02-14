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
    /// Represents an Excel table.
    /// </summary>
    internal sealed class Table : ExcelEntity
    {
        private readonly ColumnCollection columns = new ColumnCollection();

        private readonly RowCollection rows = new RowCollection();

        /// <summary>
        /// Creates a new <see cref="Table"/> object.
        /// </summary>
        public Table()
        {
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="cell">Cell name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string cell)
        {
            if (string.IsNullOrWhiteSpace(cell)) throw new ArgumentNullException("cell");

            int index = 0;
            string col = string.Empty;
            string row = string.Empty;

            var upperCell = cell.ToUpper(CultureInfo.InvariantCulture);

            for (; index < upperCell.Length && Char.IsLetter(upperCell[index]); index++) col += upperCell[index];
            for (; index < upperCell.Length && Char.IsNumber(upperCell[index]); index++) row += upperCell[index];

            return GetCell(int.Parse(row, CultureInfo.InvariantCulture) - 1, col);
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="row">Column index.</param>
        /// <param name="col">Column name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(int row, string col)
        {
            return Rows[row].GetCell(col);
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="row">Column index.</param>
        /// <param name="col">Row index.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(int row, int col)
        {
            return Rows[row].GetCell(col);
        }

        /// <summary>
        /// Returns a <see cref="Range"/>.
        /// </summary>
        /// <param name="row">Starting row.</param>
        /// <param name="col">Starting column.</param>
        /// <param name="height">Range's height.</param>
        /// <param name="width">Range's width.</param>
        /// <returns>A <see cref="Range"/>.</returns>
        public Range GetRange(int row, int col, int height, int width)
        {
            var range = new Range();

            for (var i = row; i < row + height; i++)
            {
                var rangeRow = new Row();

                for (var j = col; j < col + width; j++)
                {
                    var cell = GetCell(i, j);

                    rangeRow.Cells.Add(cell);
                }

                range.Rows.Add(rangeRow);
            }

            return range;
        }

        /// <summary>
        /// Gets the column collection.
        /// </summary>
        public ColumnCollection Columns
        {
            get { return columns; }
        }

        /// <summary>
        /// Gets the row collection.
        /// </summary>
        public RowCollection Rows
        {
            get { return rows; }
        }

        /// <summary>
        /// Gets or sets the default column width.
        /// </summary>
        public double DefaultColumnWidth
        {
            get { return double.Parse(Attributes[AttributeConstants.DefaultColumnWidthName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.DefaultColumnWidthName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the default row height.
        /// </summary>
        public long DefaultRowHeight
        {
            get { return long.Parse(Attributes[AttributeConstants.DefaultRowHeightName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.DefaultRowHeightName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the total number of columns in this table without regard for sparseness.
        /// </summary>
        public long ExpandedColumnCount
        {
            get { return long.Parse(Attributes[AttributeConstants.ExpandedColumnCountName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.ExpandedColumnCountName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the total number of rows in this table without regard for sparseness.
        /// </summary>
        public long ExpandedRowCount
        {
            get { return long.Parse(Attributes[AttributeConstants.ExpandedRowCountName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.ExpandedRowCountName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the column index that this table should be placed at.
        /// </summary>
        public long LeftCell
        {
            get { return long.Parse(Attributes[AttributeConstants.LeftCellName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.LeftCellName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the style identifier.
        /// </summary>
        public string StyleId
        {
            get { return Attributes[AttributeConstants.StyleIdName]; }
            set { Attributes[AttributeConstants.StyleIdName] = value; }
        }

        /// <summary>
        /// Gets or sets the row index that this table should be placed at.
        /// </summary>
        public long TopCell
        {
            get { return long.Parse(Attributes[AttributeConstants.TopCellName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.TopCellName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value that indicates that the source clip has full columns.
        /// </summary>
        public bool FullColumns
        {
            get
            {
                var value = Attributes[AttributeConstants.FullColumnsName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.FullColumnsName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value that indicates that the source clip has full rows.
        /// </summary>
        public bool FullRows
        {
            get
            {
                var value = Attributes[AttributeConstants.FullRowsName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.FullRowsName] = value ? "1" : "0"; }
        }

        internal override string TagName { get { return "ss:Table"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                foreach (var column in columns) xml += column.OuterXml;
                foreach (var row in rows) xml += row.OuterXml;

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml { get { return base.OuterXml + "\r\n"; } }
    }
}
