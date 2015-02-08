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
    /// Specifies the top-to-bottom alignment of text within a cell
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>
        /// The aligments applies automatically.
        /// </summary>
        Automatic,

        /// <summary>
        /// The content aligns top.
        /// </summary>
        Top,

        /// <summary>
        /// The content aligns bottom.
        /// </summary>
        Bottom,

        /// <summary>
        /// The content aligns center.
        /// </summary>
        Center,

        /// <summary>
        /// The content aligns justify.
        /// </summary>
        Justify,

        /// <summary>
        /// The content aligns distributed.
        /// </summary>
        Distributed,

        /// <summary>
        /// The content aligns justify distributed.
        /// </summary>
        JustifyDistributed,
    }
}
