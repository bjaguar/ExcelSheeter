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
    /// Specifies the fill pattern in the cell.
    /// </summary>
    public enum InteriorStylePattern
    {

        /// <summary>
        /// None pattern.
        /// </summary>
        None,

        /// <summary>
        /// Solid pattern.
        /// </summary>
        Solid,

        /// <summary>
        /// Gray75 pattern.
        /// </summary>
        Gray75,

        /// <summary>
        /// Gray50 pattern.
        /// </summary>
        Gray50,

        /// <summary>
        /// Gray25 pattern.
        /// </summary>
        Gray25,

        /// <summary>
        /// Gray125 pattern.
        /// </summary>
        Gray125,

        /// <summary>
        /// Gray0625 pattern.
        /// </summary>
        Gray0625,

        /// <summary>
        /// HorzStripe pattern.
        /// </summary>
        HorzStripe,

        /// <summary>
        /// VertStripe pattern.
        /// </summary>
        VertStripe,

        /// <summary>
        /// ReverseDiagStripe pattern.
        /// </summary>
        ReverseDiagStripe,

        /// <summary>
        /// DiagStripe pattern.
        /// </summary>
        DiagStripe,

        /// <summary>
        /// DiagCross pattern.
        /// </summary>
        DiagCross,

        /// <summary>
        /// ThickDiagCross pattern.
        /// </summary>
        ThickDiagCross,

        /// <summary>
        /// ThinHorzStripe pattern.
        /// </summary>
        ThinHorzStripe,

        /// <summary>
        /// ThinVertStripe pattern.
        /// </summary>
        ThinVertStripe,

        /// <summary>
        /// ThinReverseDiagStripe pattern.
        /// </summary>
        ThinReverseDiagStripe,

        /// <summary>
        /// ThinDiagStripe pattern.
        /// </summary>
        ThinDiagStripe,

        /// <summary>
        /// ThinHorzCross pattern.
        /// </summary>
        ThinHorzCross,

        /// <summary>
        /// ThinDiagCross pattern.
        /// </summary>
        ThinDiagCross,
    }
}
