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
    /// Represents a collection of Excel worksheets.
    /// </summary>
    public sealed class WorksheetCollection : Collection<Worksheet>
    {
        /// <summary>
        /// Returns a value that indicates if a sheet with the specified name exists in the collection.
        /// </summary>
        /// <param name="name">Sheet's name.</param>
        /// <returns>A value that indicates if a sheet with the specified name exists in the collection.</returns>
        public bool Contains(string name)
        {
            var value = this.Any(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            return value;
        }

        /// <summary>
        /// Inserts a new item.
        /// </summary>
        /// <param name="index">The item's index.</param>
        /// <param name="item">A <see cref="Worksheet"/> object.</param>
        protected override void InsertItem(int index, Worksheet item)
        {
            if (item == null) throw new ArgumentNullException("item");

            if (Contains(item.Name)) throw new ArgumentException(Resources.Exceptions.WorksheetCollection_DuplicateName);

            base.InsertItem(index, item);
        }

        /// <summary>
        /// Gets or set a sheet with the specified name.
        /// </summary>
        /// <param name="name">Sheet's name.</param>
        /// <returns>A <see cref="Worksheet"/> object.</returns>
        public Worksheet this[string name]
        {
            get
            {
                if (lastWorksheetCache != null && lastWorksheetCache.Name == name) return lastWorksheetCache;

                var sheet = this.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

                if (sheet != null) lastWorksheetCache = sheet;

                return sheet;
            }
        }

        private Worksheet lastWorksheetCache = null;
    }
}
