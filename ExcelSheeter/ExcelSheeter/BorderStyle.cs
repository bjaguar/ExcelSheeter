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
    /// Represents the border style.
    /// </summary>
    public sealed class BorderStyle : StyleEntity
    {
        /// <summary>
        /// Creates a new instance of the <see cref="BorderStyle"/> object.
        /// </summary>
        internal BorderStyle()
        {
            Position = BorderStylePosition.Bottom;
            LineStyle = BorderLineStyle.Continuous;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="BorderStyle"/> object.
        /// </summary>
        /// <param name="position">Border's position.</param>
        internal BorderStyle(BorderStylePosition position)
        {
            Position = position;
            LineStyle = BorderLineStyle.Continuous;
            Weight = 1;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="BorderStyle"/> object.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        internal BorderStyle(BorderStylePosition position, string color)
        {
            Position = position;
            Color = color;
            LineStyle = BorderLineStyle.Continuous;
            Weight = 1;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="BorderStyle"/> object.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        internal BorderStyle(BorderStylePosition position, string color, BorderLineStyle lineStyle)
        {
            Position = position;
            Color = color;
            LineStyle = lineStyle;
            Weight = 1;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="BorderStyle"/> object.
        /// </summary>
        /// <param name="position">Border's position.</param>
        /// <param name="color">Border's color.</param>
        /// <param name="lineStyle">Border's line style.</param>
        /// <param name="weight">Border's weight.</param>
        internal BorderStyle(BorderStylePosition position, string color, BorderLineStyle lineStyle, double weight)
        {
            Position = position;
            Color = color;
            LineStyle = lineStyle;
            Weight = weight;
        }

        /// <summary>
        /// Gets or sets the border's position.
        /// </summary>
        public BorderStylePosition Position
        {
            get { return (BorderStylePosition)Enum.Parse(typeof(BorderStylePosition), Attributes[AttributeConstants.PositionName]); }
            set { Attributes[AttributeConstants.PositionName] = Enum.GetName(typeof(BorderStylePosition), value); }
        }

        /// <summary>
        /// Gets or sets the border's color.
        /// </summary>
        public string Color
        {
            get { return Attributes[AttributeConstants.ColorName]; }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    Attributes.Remove(AttributeConstants.ColorName);
                }
                else
                {
                    Attributes[AttributeConstants.ColorName] = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the border's line style.
        /// </summary>
        public BorderLineStyle LineStyle
        {
            get { return (BorderLineStyle)Enum.Parse(typeof(BorderLineStyle), Attributes[AttributeConstants.LineStyleName]); }
            set { Attributes[AttributeConstants.LineStyleName] = Enum.GetName(typeof(BorderLineStyle), value); }
        }

        /// <summary>
        /// Gets or sets the border's weight.
        /// </summary>
        public double Weight
        {
            get { return double.Parse(Attributes[AttributeConstants.WeightName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.WeightName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        internal override string TagName { get { return "ss:Border"; } }
    }
}
