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
    /// Represents a cell in the Excel sheet.
    /// </summary>
    public sealed class Cell : ExcelEntity
    {
        private ExcelData data = new ExcelData();

        /// <summary>
        /// Creates a new instance of the <see cref="Cell"/> object.
        /// </summary>
        public Cell()
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="Cell"/> object.
        /// </summary>
        /// <param name="value">The cell's value.</param>
        public Cell(string value)
        {
            data.Value = value;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="Cell"/> object.
        /// </summary>
        /// <param name="value">The cell's value.</param>
        /// <param name="dataType">The cell's data type.</param>
        public Cell(object value, DataType dataType)
        {
            data.Value = value;
            data.DataType = dataType;
        }

        /// <summary>
        /// Gets or sets the <see cref="Data"/> object that represents the cell's value.
        /// </summary>
        public ExcelData Data
        {
            get { return data; }
        }

        /// <summary>
        /// Gets or sets the cell's style name.
        /// </summary>
        public string Style
        {
            get { return Attributes[AttributeConstants.StyleIdName]; }
            set
            {
                if (!string.IsNullOrEmpty(value)) Attributes[AttributeConstants.StyleIdName] = value;
                else if (Attributes.Contains(AttributeConstants.StyleIdName)) Attributes.Remove(AttributeConstants.StyleIdName);
            }
        }

        internal override string TagName { get { return "ss:Cell"; } }

        /// <summary>
        /// Gets or sets the cell's data type.
        /// </summary>
        public DataType DataType
        {
            get { return Data.DataType; }
            set { Data.DataType = value; }
        }

        /// <summary>
        /// Gets or sets the cell's value.
        /// </summary>
        public object Value
        {
            get { return Data.Value; }
            set { Data.Value = value; }
        }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                xml += data.OuterXml;

                return xml;
            }
        }

        /// <summary>
        /// Gets the outer xml.
        /// </summary>
        public override string OuterXml { get { return base.OuterXml + "\r\n"; } }
    }
}
