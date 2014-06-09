/********************************************************************************
 *   This file is part of NRtfTree Library.
 *
 *   NRtfTree Library is free software; you can redistribute it and/or modify
 *   it under the terms of the GNU Lesser General Public License as published by
 *   the Free Software Foundation; either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   NRtfTree Library is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU Lesser General Public License for more details.
 *
 *   You should have received a copy of the GNU Lesser General Public License
 *   along with this program. If not, see <http://www.gnu.org/licenses/>.
 ********************************************************************************/

/********************************************************************************
 * Library:		NRtfTree
 * Version:     v0.4
 * Date:		29/06/2013
 * Copyright:   2006-2013 Salvador Gomez
 * Home Page:	http://www.sgoliver.net
 * GitHub:	    https://github.com/sgolivernet/nrtftree
 * Class:		RtfFont
 * Description:	Tabla de Fuentes de un documento RTF.
 * ******************************************************************************/

using System;
using System.Collections.Generic;
using System.Collections;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// An item of font table.
        /// </summary>
        public class RtfFont
        {
            /// <summary>
            /// Font name
            /// </summary>
            public string Name{get; private set;}

            /// <summary>
            /// Code page
            /// </summary>
            public int CodePage{get; private set;}

            /// <summary>
            /// 
            /// </summary>
            /// <param name="name"></param>
            /// <param name="codePage"></param>
            public RtfFont(string name, int codePage)
            {
                this.Name = name;
                this.CodePage = codePage;
            }

        }

    }
}
