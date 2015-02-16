﻿/*
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
using System.Security;

namespace ExcelSheeter
{
    /// <summary>
    /// Represents an Excel attribute.
    /// </summary>
    internal sealed class ExcelAttribute
    {
        private readonly string _key = string.Empty;

        private string _value = string.Empty;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="key">The attribute's key.</param>
        public ExcelAttribute(string key)
        {
            _key = key;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="key">Attribute's key.</param>
        /// <param name="value">Attribute's value.</param>
        public ExcelAttribute(string key, string value)
        {
            _key = key;
            Value = value;
        }

        /// <summary>
        /// Gets or sets the attribute key.
        /// </summary>
        public string Key
        {
            get { return _key; }
        }

        /// <summary>
        /// Gets or sets the attribute value.
        /// </summary>
        public string Value
        {
            get { return _value; }
            set
            {
                var encodedValue = SecurityElement.Escape(value);
                
                _value = encodedValue;
            }
        }
    }
}
