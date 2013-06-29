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
 * Class:		RtfDocumentFormat
 * Description:	Representa un formato de documento.
 * ******************************************************************************/

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Representa un formato de documento.
        /// </summary>
        public class RtfDocumentFormat
        {
            #region Atributos Privados

            /// <summary>
            /// Margen izquierdo (en cm.)
            /// </summary>
            private float marginl = 2;

            /// <summary>
            /// Margen derecho (en cm.)
            /// </summary>
            private float marginr = 2;

            /// <summary>
            /// Margen superior (en cm.)
            /// </summary>
            private float margint = 2;

            /// <summary>
            /// Margen inferior (en cm.)
            /// </summary>
            private float marginb = 2;

            #endregion

            #region Propiedades

            /// <summary>
            /// Margen Izquierdo en centímetros.
            /// </summary>
            public float MarginL
            {
                get { return marginl; }
                set { marginl = value; }
            }

            /// <summary>
            /// Margen Derecho en centímetros.
            /// </summary>
            public float MarginR
            {
                get { return marginr; }
                set { marginr = value; }
            }

            /// <summary>
            /// Margen Superior en centímetros.
            /// </summary>
            public float MarginT
            {
                get { return margint; }
                set { margint = value; }
            }

            /// <summary>
            /// Margen Inferior en centímetros.
            /// </summary>
            public float MarginB
            {
                get { return marginb; }
                set { marginb = value; }
            }

            #endregion
        }
    }
}
