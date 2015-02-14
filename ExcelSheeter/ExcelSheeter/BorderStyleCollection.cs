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
    /// Represents a collection of border styles.
    /// </summary>
    public sealed class BorderStyleCollection : IEnumerable<BorderStyle>
    {
        private readonly Collection<BorderStyle> items = new Collection<BorderStyle>();

        internal BorderStyleCollection()
        {
        }

        internal bool Remove(BorderStylePosition position)
        {
            var item = items.FirstOrDefault(x => x.Position == position);

            if (item != null)
            {
                return items.Remove(item);
            }

            return false;
        }

        internal void RemoveAll()
        {
            Remove(BorderStylePosition.Bottom);
            Remove(BorderStylePosition.DiagonalLeft);
            Remove(BorderStylePosition.DiagonalRight);
            Remove(BorderStylePosition.Left);
            Remove(BorderStylePosition.Right);
            Remove(BorderStylePosition.Top);
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <remarks>
        /// <para>The border applies to bottom, top, left and right.</para>
        /// </remarks>
        /// <param name="color">Border's color.</param>
        public void Add(string color)
        {
            RemoveAll();

            items.Add(new BorderStyle(BorderStylePosition.Bottom, color));
            items.Add(new BorderStyle(BorderStylePosition.Left, color));
            items.Add(new BorderStyle(BorderStylePosition.Right, color));
            items.Add(new BorderStyle(BorderStylePosition.Top, color));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <remarks>
        /// <para>The border applies to bottom, top, left and right.</para>
        /// </remarks>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        public void Add(string color, BorderLineStyle lineStyle)
        {
            RemoveAll();

            items.Add(new BorderStyle(BorderStylePosition.Bottom, color, lineStyle));
            items.Add(new BorderStyle(BorderStylePosition.Left, color, lineStyle));
            items.Add(new BorderStyle(BorderStylePosition.Right, color, lineStyle));
            items.Add(new BorderStyle(BorderStylePosition.Top, color, lineStyle));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <remarks>
        /// <para>The border applies to bottom, top, left and right.</para>
        /// </remarks>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        /// <param name="weight">Border's weight.</param>
        public void Add(string color, BorderLineStyle lineStyle, double weight)
        {
            RemoveAll();

            items.Add(new BorderStyle(BorderStylePosition.Bottom, color, lineStyle, weight));
            items.Add(new BorderStyle(BorderStylePosition.Left, color, lineStyle, weight));
            items.Add(new BorderStyle(BorderStylePosition.Right, color, lineStyle, weight));
            items.Add(new BorderStyle(BorderStylePosition.Top, color, lineStyle, weight));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        public void Add(BorderStylePosition position, string color)
        {
            Remove(position);

            items.Add(new BorderStyle(position, color));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        public void Add(BorderStylePosition position, string color, BorderLineStyle lineStyle)
        {
            Remove(position);

            items.Add(new BorderStyle(position, color, lineStyle));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        /// <param name="weight">Border's weight.</param>
        public void Add(BorderStylePosition position, string color, BorderLineStyle lineStyle, double weight)
        {
            Remove(position);

            items.Add(new BorderStyle(position, color, lineStyle, weight));
        }

        /// <summary>
        /// Gets an enumerator.
        /// </summary>
        /// <returns>An <see cref="IEnumerator"/> object.</returns>
        public IEnumerator<BorderStyle> GetEnumerator()
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
        /// Gets the specified border.
        /// </summary>
        /// <param name="position">The border's position.</param>
        /// <returns>A <see cref="BorderStyle"/> object.</returns>
        public BorderStyle this[BorderStylePosition position]
        {
            get
            {
                var border = items.FirstOrDefault(x => x.Position == position);

                if (border == null) throw new KeyNotFoundException();

                return border;
            }
        }

        /// <summary>
        /// Gets the item count.
        /// </summary>
        public int Count { get { return items.Count; } }
    }
}
