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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents the base class for all the main entities of the ExcelSheeter namespace.
    /// </summary>
    public abstract class ExcelEntity
    {
        private readonly ExcelAttributeCollection attributes = new ExcelAttributeCollection();

        /// <summary>
        /// Gets the tag name of the entity.
        /// </summary>
        internal abstract string TagName { get; }

        /// <summary>
        /// Gets the attribute collection.
        /// </summary>
        internal ExcelAttributeCollection Attributes { get { return attributes; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public virtual string InnerXml { get { return string.Empty; } }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public virtual string OuterXml
        {
            get
            {
                var xml = string.Empty;
                var innerXml = InnerXml;

                var attributeString = string.Empty;
                foreach (var attribute in Attributes) attributeString += string.Format(CultureInfo.InvariantCulture, @" {0}=""{1}""", attribute.Key, attribute.Value);

                if (string.IsNullOrWhiteSpace(innerXml))
                {
                    xml += string.Format(CultureInfo.InvariantCulture, "<{0}{1} />", TagName, attributeString);
                }
                else
                {
                    xml += string.Format(CultureInfo.InvariantCulture, "<{0}{1}>{2}</{0}>", TagName, attributeString, innerXml);
                }

                return xml;
            }
        }
    }
}
