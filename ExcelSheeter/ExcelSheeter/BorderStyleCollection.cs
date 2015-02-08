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
    /// Represents a collection of border styles.
    /// </summary>
    public sealed class BorderStyleCollection : Collection<BorderStyle>
    {
        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        public void Add(BorderStylePosition position)
        {
            Add(new BorderStyle(position));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        public void Add(BorderStylePosition position, string color)
        {
            Add(new BorderStyle(position, color));
        }

        /// <summary>
        /// Adds a new border style.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        public void Add(BorderStylePosition position, string color, BorderLineStyle lineStyle)
        {
            Add(new BorderStyle(position, color, lineStyle));
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
            Add(new BorderStyle(position, color, lineStyle, weight));
        }
    }
}
