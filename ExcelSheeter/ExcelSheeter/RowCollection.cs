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
    /// Represents a collection of <see cref="Row"/> objects.
    /// </summary>
    public sealed class RowCollection : Collection<Row>
    {
        /// <summary>
        /// Gets or sets a <see cref="Row"/> object.
        /// </summary>
        /// <param name="index">Index of the object.</param>
        /// <returns>A <see cref="Row"/> object.</returns>
        public new Row this[int index]
        {
            get
            {
                return (Row)base[index];
            }
            set
            {
                base[index] = value;
            }
        }
    }
}
