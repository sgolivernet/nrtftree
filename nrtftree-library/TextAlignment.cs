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
 * Version:     v0.3.0
 * Date:		01/05/2009
 * Copyright:   2006-2009 Salvador Gomez
 * E-mail:      sgoliver.net@gmail.com
 * Home Page:	http://www.sgoliver.net
 * SF Project:	http://nrtftree.sourceforge.net
 *				http://sourceforge.net/projects/nrtftree
 * Class:		TextAlignment
 * Description:	Tipos de alineación de texto.
 * ******************************************************************************/

using System;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Tipos de alineación de texto.
        /// </summary>
        public enum TextAlignment
        {
            /// <summary>
            /// Texto alineado a la izquierda.
            /// </summary>
            Left = 0,
            /// <summary>
            /// Texto alineado a la derecha.
            /// </summary>
            Right = 1,
            /// <summary>
            /// Texto centrado.
            /// </summary>
            Centered = 2,
            /// <summary>
            /// Texto justificado.
            /// </summary>
            Justified = 3
        }
    }
}
