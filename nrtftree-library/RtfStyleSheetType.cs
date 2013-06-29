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
 * Class:		RtfStyleSheetType
 * Description:	Tipos de hojas de estilo de documento RTF.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Tipos de hojas de estilo de un documento RTF.
        /// </summary>
        public enum RtfStyleSheetType
        {
            /// <summary>
            /// Hoja de estilo sin definir.
            /// </summary>
            None = 0,
            /// <summary>
            /// Hoja de estilo de caracter.
            /// </summary>
            Character = 1,
            /// <summary>
            /// Hoja de estilo de párrafo.
            /// </summary>
            Paragraph = 2,
            /// <summary>
            /// Hoja de estilo de sección.
            /// </summary>
            Section = 3,
            /// <summary>
            /// Hoja de estilo de tabla.
            /// </summary>
            Table = 4
        }
    }
}