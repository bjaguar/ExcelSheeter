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
    /// Represents an Excel workbook.
    /// </summary>
    public sealed class Workbook : ExcelEntity
    {
        private readonly ExcelStyles styles = new ExcelStyles();

        private readonly WorksheetCollection sheets = new WorksheetCollection();

        /// <summary>
        /// Creates a new instance of the <see cref="Workbook"/> object.
        /// </summary>
        public Workbook()
        {
            Attributes["xmlns"] = "urn:schemas-microsoft-com:office:spreadsheet";
            Attributes["xmlns:o"] = "urn:schemas-microsoft-com:office:office";
            Attributes["xmlns:x"] = "urn:schemas-microsoft-com:office:excel";
            Attributes["xmlns:c"] = "urn:schemas-microsoft-com:office:component:spreadsheet";
            Attributes["xmlns:ss"] = "urn:schemas-microsoft-com:office:spreadsheet";
            Attributes["xmlns:html"] = "http://www.w3.org/TR/REC-html40";
        }

        /// <summary>
        /// Returns a worksheet given it's name.
        /// </summary>
        /// <param name="name">Worksheet's name.</param>
        /// <returns>A <see cref="Worksheet"/> object.</returns>
        public Worksheet GetWorksheet(string name)
        {
            Worksheet sheet;

            var value = TryGetWorksheet(name, out sheet);

            if (!value) throw new InvalidOperationException(Resources.Exceptions.WorksheetCollection_ItemNotFound);

            return sheet;
        }

        /// <summary>
        /// Tries to return one worksheet.
        /// </summary>
        /// <param name="name">Worksheet's name.</param>
        /// <param name="sheet">A <see cref="Worksheet"/> output parameter.</param>
        /// <returns>A value indicating if the worksheet was returned.</returns>
        public bool TryGetWorksheet(string name, out Worksheet sheet)
        {
            sheet = Sheets.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            return sheet != null;
        }

        /// <summary>
        /// Returns a <see cref="Style"/> object given it's identifier.
        /// </summary>
        /// <param name="id">Style's identifier.</param>
        /// <returns>A <see cref="Style"/> object.</returns>
        public Style GetStyle(string id)
        {
            Style style;

            var value = TryGetStyle(id, out style);

            if (!value) throw new InvalidOperationException(Resources.Exceptions.StyleCollection_ItemNotFound);

            return style;
        }

        /// <summary>
        /// Tries to obtain a <see cref="Style"/> object given it's identifier.
        /// </summary>
        /// <param name="id">Style's identifier.</param>
        /// <param name="style">A <see cref="Style"/> object.</param>
        /// <returns>A value indicating if the style was found.</returns>
        public bool TryGetStyle(string id, out Style style)
        {
            style = Styles.FirstOrDefault(x => x.Id.Equals(id, StringComparison.OrdinalIgnoreCase));

            return style != null;
        }

        /// <summary>
        /// Adds a new style.
        /// </summary>
        /// <param name="item">A <see cref="Style"/> object.</param>
        internal void AddStyle(Style item)
        {
            Styles.Add(item);
        }

        /// <summary>
        /// Adds a new style.
        /// </summary>
        /// <param name="id">Style's identifier.</param>
        /// <returns>A <see cref="Style"/> object.</returns>
        public Style AddStyle(string id)
        {
            var style = new Style(id);

            AddStyle(style);

            return style;
        }

        /// <summary>
        /// Removes one style from the book.
        /// </summary>
        /// <param name="item">A <see cref="Style"/> object.</param>
        /// <returns>A value indicating if the style was removed.</returns>
        public bool RemoveStyle(Style item)
        {
            var value = Styles.Remove(item);

            return value;
        }

        /// <summary>
        /// Removes one style from the book.
        /// </summary>
        /// <param name="id">Style's identifier.</param>
        /// <returns>A value indicating if the style was removed.</returns>
        public bool RemoveStyle(string id)
        {
            Style style;

            var value = TryGetStyle(id, out style);

            if (value)
            {
                return RemoveStyle(style);
            }

            return false;
        }

        /// <summary>
        /// Removes all styles of the book.
        /// </summary>
        public void ClearStyles()
        {
            Styles.Clear();
        }

        /// <summary>
        /// Adds a new sheet to the book.
        /// </summary>
        /// <param name="sheet">A <see cref="Worksheet"/> object.</param>
        internal void AddSheet(Worksheet sheet)
        {
            Sheets.Add(sheet);
        }
        
        /// <summary>
        /// Adds a new sheet to the book.
        /// </summary>
        /// <param name="name">Worksheet's name.</param>
        /// <returns>A <see cref="Worksheet"/> object.</returns>
        public Worksheet AddSheet(string name)
        {
            var sheet = new Worksheet(name);

            AddSheet(sheet);

            return sheet;
        }

        /// <summary>
        /// Removes one sheet of the book.
        /// </summary>
        /// <param name="sheet">A <see cref="Worksheet"/> object.</param>
        /// <returns>A value indicating if the sheet was removed.</returns>
        public bool RemoveSheet(Worksheet sheet)
        {
            var value = Sheets.Remove(sheet);

            return value;
        }

        /// <summary>
        /// Removes one sheet of the book.
        /// </summary>
        /// <param name="name">The worksheet's name.</param>
        /// <returns>A value indicating if the sheet was removed.</returns>
        public bool RemoveSheet(string name)
        {
            Worksheet sheet;
            var value = TryGetWorksheet(name, out sheet);

            if (value)
            {
                return RemoveSheet(sheet);
            }

            return value;
        }

        /// <summary>
        /// Removes all the sheets in the book.
        /// </summary>
        public void ClearSheets()
        {
            Sheets.Clear();
        }

        /// <summary>
        /// Gets the styles of this book.
        /// </summary>
        internal StyleCollection Styles { get { return styles.Styles; } }

        /// <summary>
        /// Gets the sheets of this book.
        /// </summary>
        internal WorksheetCollection Sheets { get { return sheets; } }

        internal override string TagName { get { return "ss:Workbook"; } }

        /// <summary>
        /// Gets the inner xml.
        /// </summary>
        public override string InnerXml
        {
            get
            {
                var xml = string.Empty;

                xml += @"<DocumentProperties xmlns=""urn:schemas-microsoft-com:office:office""></DocumentProperties><OfficeDocumentSettings xmlns=""urn:schemas-microsoft-com:office:office""><AllowPNG/><RemovePersonalInformation/></OfficeDocumentSettings><ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel""><WindowHeight>8010</WindowHeight><WindowWidth>14805</WindowWidth><WindowTopX>240</WindowTopX><WindowTopY>105</WindowTopY><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook>";

                xml += Styles.OuterXml;
                
                foreach (var sheet in Sheets) xml += sheet.OuterXml;

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
                var xml = string.Empty;

                xml += @"<?xml version=""1.0""?>" + "\r\n";
                xml += @"<?mso-application progid=""Excel.Sheet""?>" + "\r\n";
               
                xml += base.OuterXml;

                return xml + "\r\n";
            }
        }
    }
}
