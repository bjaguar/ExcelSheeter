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
    /// Represents the protection style.
    /// </summary>
    public sealed class ProtectionStyle : StyleEntity
    {
        /// <summary>
        /// Gets or sets a value indicating if the protection is activated.
        /// </summary>
        public bool Protected
        {
            get
            {
                var value = Attributes[AttributeConstants.ProtectedName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.ProtectedName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether or not this cell's formula should be hidden when worksheet protection is enabled.
        /// </summary>
        public bool HideFormula
        {
            get
            {
                var value = Attributes[AttributeConstants.HideFormulaName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.HideFormulaName] = value ? "1" : "0"; }
        }

        internal override string TagName { get { return "ss:Protection"; } }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml
        {
            get
            {
                if (Attributes.Count > 0) return base.OuterXml;

                return string.Empty;
            }
        }
    }
}
