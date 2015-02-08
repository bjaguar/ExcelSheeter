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
    /// Represents an Excel range.
    /// </summary>
    public sealed class Range
    {
        private RowCollection rows = new RowCollection();

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>A <see cref="Cell"/> object</returns>
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
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        /// <returns>A <see cref="Cell"/> object</returns>
        public Cell GetCell(int row, string column)
        {
            return Rows[row].GetCell(column);
        }

        /// <summary>
        /// Returns a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        /// <returns>A <see cref="Cell"/> object</returns>
        public Cell GetCell(int row, int column)
        {
            return Rows[row].GetCell(column);
        }

        /// <summary>
        /// Gets the rows collection.
        /// </summary>
        public RowCollection Rows { get { return rows; } }
    }
}
