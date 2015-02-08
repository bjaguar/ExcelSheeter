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
            var sheet = Sheets.Single(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            return sheet;
        }

        /// <summary>
        /// Returns the specified cell.
        /// </summary>
        /// <param name="worksheetName">Worksheet's name.</param>
        /// <param name="cellName">Cell's name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string worksheetName, string cellName)
        {
            var worksheet = GetWorksheet(worksheetName);

            var cell = worksheet.GetCell(cellName);

            return cell;
        }

        /// <summary>
        /// Returns the specified cell.
        /// </summary>
        /// <param name="worksheetName">Worksheet's name.</param>
        /// <param name="row">Row's index.</param>
        /// <param name="col">Col's name.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string worksheetName, int row, string col)
        {
            var worksheet = GetWorksheet(worksheetName);

            var cell = worksheet.GetCell(row, col);

            return cell;
        }

        /// <summary>
        /// Returns the specified cell.
        /// </summary>
        /// <param name="worksheetName">Worksheet's name.</param>
        /// <param name="row">Row's index.</param>
        /// <param name="col">Col's index.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell GetCell(string worksheetName, int row, int col)
        {
            var worksheet = GetWorksheet(worksheetName);

            var cell = worksheet.GetCell(row, col);

            return cell;
        }

        /// <summary>
        /// Gets the styles of this book.
        /// </summary>
        public StyleCollection Styles { get { return styles.Styles; } }

        /// <summary>
        /// Gets the sheets of this book.
        /// </summary>
        public WorksheetCollection Sheets { get { return sheets; } }

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
