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
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents a list of Excel attributes.
    /// </summary>
    internal sealed class ExcelAttributeCollection : Collection<ExcelAttribute>
    {
        /// <summary>
        /// Adds a new attribute to the list.
        /// </summary>
        /// <param name="key">Key of the attribute.</param>
        /// <param name="value">Value of the attribute.</param>
        public void Add(string key, string value)
        {
            var item = new ExcelAttribute(key, value);

            Add(item);
        }

        /// <summary>
        /// Inserts a new item.
        /// </summary>
        /// <param name="index">Item's index.</param>
        /// <param name="item">A <see cref="ExcelAttribute"/> object.</param>
        protected override void InsertItem(int index, ExcelAttribute item)
        {
            if (item == null) throw new ArgumentNullException("item");

            if (Contains(item.Key)) throw new ArgumentException(Resources.Exceptions.ExcelAttributeCollection_DuplicateKey);

            base.InsertItem(index, item);
        }

        /// <summary>
        /// Replaces the element at the specified index.
        /// </summary>
        /// <param name="index">Item's index.</param>
        /// <param name="item">A <see cref="ExcelAttribute"/> object.</param>
        protected override void SetItem(int index, ExcelAttribute item)
        {
            if (item == null) throw new ArgumentNullException("item");

            if (Contains(item.Key)) throw new ArgumentException(Resources.Exceptions.ExcelAttributeCollection_DuplicateKey);

            base.SetItem(index, item);
        }

        /// <summary>
        /// Returns a value indicating if a key exists in the list.
        /// </summary>
        /// <param name="key">Key of the attribute.</param>
        /// <returns>A value indicating if the key exists in the list.</returns>
        public bool Contains(string key)
        {
            var value = this.Any(x => x.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

            return value;
        }

        /// <summary>
        /// Removes an item from the list giving it's key.
        /// </summary>
        /// <param name="key">Key of the item to remove.</param>
        /// <returns>A value indicating if the item was deleted.</returns>
        public bool Remove(string key)
        {
            var item = this.FirstOrDefault(x => x.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

            if (item != null)
            {
                Remove(item);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Gets or sets the value of the item that contains the specified key.
        /// </summary>
        /// <param name="key">Key of the item.</param>
        /// <returns>The value of the matching item.</returns>
        public string this[string key]
        {
            get
            {
                var item = this.Single(x => x.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

                return item.Value;
            }
            set
            {
                var item = this.FirstOrDefault(x => x.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

                if (item != null)
                {
                    item.Value = value;
                }
                else
                {
                    Add(key, value);
                }
            }
        }
    }
}
