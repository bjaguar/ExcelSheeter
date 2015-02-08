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
    /// Represents a list of Excel attributes.
    /// </summary>
    public sealed class ExcelAttributeCollection : List<ExcelAttribute>
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
        /// Returns a value indicating if a key exists in the list.
        /// </summary>
        /// <param name="key">Key of the attribute.</param>
        /// <returns>A value indicating if the key exists in the list.</returns>
        public bool Contains(string key)
        {
            var value = this.Any(x => x.Key == key);

            return value;
        }

        /// <summary>
        /// Removes an item from the list giving it's key.
        /// </summary>
        /// <param name="key">Key of the item to remove.</param>
        /// <returns>A value indicating if the item was deleted.</returns>
        public bool Remove(string key)
        {
            var item = this.FirstOrDefault(x => x.Key == key);

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
                var item = this.Single(x => x.Key == key);

                return item.Value;
            }
            set
            {
                var item = this.FirstOrDefault(x => x.Key == key);

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
