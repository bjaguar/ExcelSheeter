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
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents a style collection.
    /// </summary>
    public sealed class StyleCollection : ExcelEntity, ICollection<Style>
    {
        private Collection<Style> items = new Collection<Style>();

        /// <summary>
        /// Adds a new style to the collection.
        /// </summary>
        /// <param name="item">A <see cref="Style"/> object.</param>
        public void Add(Style item)
        {
            items.Add(item);
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            items.Clear();
        }

        /// <summary>
        /// Returns a value indicating if the item exists in the collection.
        /// </summary>
        /// <param name="item">A <see cref="Style"/> object.</param>
        /// <returns>A value indicating if the item exists in the collection.</returns>
        public bool Contains(Style item)
        {
            var value = items.Contains(item);

            return value;
        }

        /// <summary>
        /// Copies the array, starting in the specified arrayIndex.
        /// </summary>
        /// <param name="array">Array to copy.</param>
        /// <param name="arrayIndex">Start index.</param>
        public void CopyTo(Style[] array, int arrayIndex)
        {
            items.CopyTo(array, arrayIndex);
        }

        /// <summary>
        /// Gets the item count.
        /// </summary>
        public int Count { get { return items.Count; } }

        /// <summary>
        /// Returns a value indicating if the collection is read only.
        /// </summary>
        public bool IsReadOnly { get { return false; } }

        /// <summary>
        /// Removes an item from the collection.
        /// </summary>
        /// <param name="item">Item to remove.</param>
        /// <returns>A value indicating if the item was deleted.</returns>
        public bool Remove(Style item)
        {
            var value = items.Remove(item);

            return value;
        }

        /// <summary>
        /// Returns an enumerator.
        /// </summary>
        /// <returns>An enumerator.</returns>
        public IEnumerator<Style> GetEnumerator()
        {
            return items.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator.
        /// </summary>
        /// <returns>An enumerator.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return items.GetEnumerator();
        }

        internal override string TagName { get { return "ss:Styles"; } }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml
        {
            get
            {
                if (items.Count > 0) return base.OuterXml;

                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                foreach (var item in items) xml += item.OuterXml;

                return xml;
            }
        }
    }
}
