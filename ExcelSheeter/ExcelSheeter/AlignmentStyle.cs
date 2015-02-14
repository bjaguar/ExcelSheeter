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
    /// Represents an Excel alignment style.
    /// </summary>
    public sealed class AlignmentStyle : StyleEntity
    {
        internal AlignmentStyle()
        {
        }

        /// <summary>
        /// Gets or sts the horizontal alignment.
        /// </summary>
        public HorizontalAlignment Horizontal
        {
            get { return (HorizontalAlignment)Enum.Parse(typeof(HorizontalAlignment), Attributes[AttributeConstants.HorizontalName]); }
            set { Attributes[AttributeConstants.HorizontalName] = Enum.GetName(typeof(HorizontalAlignment), value); }
        }

        /// <summary>
        /// Gets or sets the indent.
        /// </summary>
        public long Indent
        {
            get { return long.Parse(Attributes[AttributeConstants.IndentName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.IndentName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets the reading order.
        /// </summary>
        public ReadingOrder ReadingOrder
        {
            get { return (ReadingOrder)Enum.Parse(typeof(ReadingOrder), Attributes[AttributeConstants.ReadingOrderName]); }
            set { Attributes[AttributeConstants.ReadingOrderName] = Enum.GetName(typeof(ReadingOrder), value); }
        }

        /// <summary>
        /// Gets or sets the rotation angle.
        /// </summary>
        public double Rotate
        {
            get { return double.Parse(Attributes[AttributeConstants.RotateName], CultureInfo.InvariantCulture); }
            set { Attributes[AttributeConstants.RotateName] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Gets or sets a value indicating if the text shrink to fit.
        /// </summary>
        public bool ShrinkToFit
        {
            get
            {
                var value = Attributes[AttributeConstants.ShrinkToFitName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.ShrinkToFitName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public VerticalAlignment Vertical
        {
            get { return (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), Attributes[AttributeConstants.VerticalName]); }
            set { Attributes[AttributeConstants.VerticalName] = Enum.GetName(typeof(VerticalAlignment), value); }
        }

        /// <summary>
        /// Gets or sets a value indicating if the text is vertical.
        /// </summary>
        public bool VerticalText
        {
            get
            {
                var value = Attributes[AttributeConstants.VerticalTextName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.VerticalTextName] = value ? "1" : "0"; }
        }

        /// <summary>
        /// Gets or sets a value indicating if the text wraps.
        /// </summary>
        public bool WrapText
        {
            get
            {
                var value = Attributes[AttributeConstants.WrapTextName];

                return value == "1";
            }
            set { Attributes[AttributeConstants.WrapTextName] = value ? "1" : "0"; }
        }

        internal override string TagName { get { return "ss:Alignment"; } }

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
