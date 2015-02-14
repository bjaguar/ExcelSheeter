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
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents a collection of Excel cells.
    /// </summary>
    public sealed class CellCollection : IEnumerable<Cell>
    {
        private readonly Collection<Cell> items = new Collection<Cell>();

        internal CellCollection()
        {
        }

        internal void Add(Cell item)
        {
            items.Add(item);
        }

        /// <summary>
        /// Gets an enumerator.
        /// </summary>
        /// <returns>An <see cref="IEnumerator"/> object.</returns>
        public IEnumerator<Cell> GetEnumerator()
        {
            return items.GetEnumerator();
        }

        /// <summary>
        /// Gets an enumerator.
        /// </summary>
        /// <returns>An <see cref="IEnumerator"/> object.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return items.GetEnumerator();
        }

        /// <summary>
        /// Gets a cell in the collection given it's index.
        /// </summary>
        /// <param name="index">Index of the cell.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell this[int index]
        {
            get
            {
                while (index >= Count)
                {
                    Add(new Cell());
                }

                return items[index];
            }
        }

        /// <summary>
        /// Gets the cell count.
        /// </summary>
        public int Count { get { return items.Count; } }
    }
}
