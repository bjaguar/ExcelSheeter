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
    /// Represents the cell's borders.
    /// </summary>
    public sealed class BordersStyle : StyleEntity
    {
        private readonly BorderStyleCollection borders = new BorderStyleCollection();

        internal BordersStyle()
        {
        }

        /// <summary>
        /// Gets the border collection.
        /// </summary>
        public BorderStyleCollection Borders
        {
            get { return borders; }
        }

        internal override string TagName { get { return "ss:Borders"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                foreach (var border in borders) xml += border.OuterXml;

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml
        {
            get
            {
                if (borders.Count > 0) return base.OuterXml;

                return string.Empty;
            }
        }
    }
}
