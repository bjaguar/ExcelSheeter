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
    /// Represents an Excel worksheet.
    /// </summary>
    public sealed class Worksheet : ExcelEntity
    {
        private readonly Table table = new Table();

        /// <summary>
        /// Creates a new <see cref="Worksheet"/> object.
        /// </summary>
        /// <param name="name">The sheet name.</param>
        internal Worksheet(string name)
        {
            Attributes[AttributeConstants.Name] = name;
        }

        /// <summary>
        /// Gets the table asociated to this sheet.
        /// </summary>
        internal Table Table
        {
            get { return table; }
        }

        /// <summary>
        /// Gets or sets the sheet name.
        /// </summary>
        public string Name
        {
            get { return Attributes[AttributeConstants.Name]; }
        }

        internal override string TagName { get { return "ss:Worksheet"; } }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="cell">Cell name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string cell)
        {
            return Table.GetCell(cell);
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="row">Column index.</param>
        /// <param name="col">Column name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(int row, string col)
        {
            return Table.GetCell(row, col);
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="row">Column index.</param>
        /// <param name="col">Row index.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(int row, int col)
        {
            return Table.GetCell(row, col);
        }

        /// <summary>
        /// Gets the cell with the specified name.
        /// </summary>
        /// <param name="name">Cell's name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell this[string name] { get { return GetCell(name); } }

        /// <summary>
        /// Gets the column collection.
        /// </summary>
        public ColumnCollection Columns { get { return Table.Columns; } }

        /// <summary>
        /// Gets the row collection.
        /// </summary>
        public RowCollection Rows { get { return Table.Rows; } }

        /// <summary>
        /// Gets or sets the default column width.
        /// </summary>
        public double DefaultColumnWidth
        {
            get { return Table.DefaultColumnWidth; }
            set { Table.DefaultColumnWidth = value; }
        }

        /// <summary>
        /// Gets or sets the default row height.
        /// </summary>
        public long DefaultRowHeight
        {
            get { return Table.DefaultRowHeight; }
            set { Table.DefaultRowHeight = value; }
        }

        /// <summary>
        /// Gets or sets the total number of columns in this table without regard for sparseness.
        /// </summary>
        public long ExpandedColumnCount
        {
            get { return Table.ExpandedColumnCount; }
            set { Table.ExpandedColumnCount = value; }
        }

        /// <summary>
        /// Gets or sets the total number of rows in this table without regard for sparseness.
        /// </summary>
        public long ExpandedRowCount
        {
            get { return Table.ExpandedRowCount; }
            set { Table.ExpandedRowCount = value; }
        }

        /// <summary>
        /// Gets or sets the column index that this table should be placed at.
        /// </summary>
        public long LeftCell
        {
            get { return Table.LeftCell; }
            set { Table.LeftCell = value; }
        }

        /// <summary>
        /// Gets or sets the style identifier.
        /// </summary>
        public string StyleId
        {
            get { return Table.StyleId; }
            set { Table.StyleId = value; }
        }

        /// <summary>
        /// Gets or sets the row index that this table should be placed at.
        /// </summary>
        public long TopCell
        {
            get { return Table.TopCell; }
            set { Table.TopCell = value; }
        }

        /// <summary>
        /// Gets or sets a value that indicates that the source clip has full columns.
        /// </summary>
        public bool FullColumns
        {
            get { return Table.FullColumns; }
            set { Table.FullColumns = value; }
        }

        /// <summary>
        /// Gets or sets a value that indicates that the source clip has full rows.
        /// </summary>
        public bool FullRows
        {
            get { return Table.FullRows; }
            set { Table.FullRows = value; }
        }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                xml += table.OuterXml;

                xml += "<WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'><PageSetup><Layout x:Orientation='Landscape'/><Header x:Margin='0'/><Footer x:Margin='0'/></PageSetup><FitToPage/></WorksheetOptions>";

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml { get { return base.OuterXml + "\r\n"; } }
    }
}
