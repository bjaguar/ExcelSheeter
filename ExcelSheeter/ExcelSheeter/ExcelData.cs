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
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents the value of one cell specifying the data type.
    /// </summary>
    internal sealed class ExcelData : ExcelEntity
    {
        private object _value = string.Empty;

        /// <summary>
        /// Creates a new instance of the <see cref="ExcelData"/> object.
        /// </summary>
        /// <remarks>
        /// <para>By default takes the empty string value.</para>
        /// </remarks>
        public ExcelData()
        {
            DataType = DataType.String;
            _value = string.Empty;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="ExcelData"/> object.
        /// </summary>
        /// <remarks>
        /// <para>By default takes the string data type.</para>
        /// </remarks>
        /// <param name="value">The cell's value.</param>
        public ExcelData(string value)
        {
            DataType = DataType.String;
            Value = value;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="ExcelData"/> object.
        /// </summary>
        /// <param name="dataType">The cell's data type.</param>
        /// <param name="value">The cell's value.</param>
        public ExcelData(DataType dataType, object value)
        {
            DataType = dataType;
            Value = value;
        }

        /// <summary>
        /// Gets or sets the data type of the cell.
        /// </summary>
        public DataType DataType
        {
            get { return (DataType)Enum.Parse(typeof(DataType), Attributes[AttributeConstants.TypeName]); }
            set { Attributes[AttributeConstants.TypeName] = Enum.GetName(typeof(DataType), value); }
        }

        /// <summary>
        /// Gets or sets the value of the cell.
        /// </summary>
        public object Value
        {
            get { return _value; }
            set
            {
                var stringValue = value as string;

                if (stringValue != null)
                {
                    var encodedValue = SecurityElement.Escape(stringValue);

                    _value = encodedValue;
                }
                else
                {
                    _value = value;
                }
            }
        }

        internal override string TagName { get { return "ss:Data"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                xml += _value.ToString();

                return xml;
            }
        }
    }
}
