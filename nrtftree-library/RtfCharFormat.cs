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
 * Class:		RtfCharFormat
 * Description:	Representa un formato de texto.
 * ******************************************************************************/

using System.Drawing;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Representa un formato de texto.
        /// </summary>
        public class RtfCharFormat
        {
            #region Atributos Privados

            /// <summary>
            /// Negrita.
            /// </summary>
            private bool bold = false;

            /// <summary>
            /// Cursiva.
            /// </summary>
            private bool italic = false;

            /// <summary>
            /// Subrayado.
            /// </summary>
            private bool underline = false;

            /// <summary>
            /// Nombre de la fuente.
            /// </summary>
            private string font = "Arial";

            /// <summary>
            /// Tamaño de la fuente.
            /// </summary>
            private int size = 10;

            /// <summary>
            /// Color de la fuente.
            /// </summary>
            private Color color = Color.Black;

            #endregion

            #region Propiedades

            /// <summary>
            /// Fuente negrita.
            /// </summary>
            public bool Bold
            {
                get { return bold; }
                set { bold = value; }
            }

            /// <summary>
            /// Fuente cursiva.
            /// </summary>
            public bool Italic
            {
                get { return italic; }
                set { italic = value; }
            }

            /// <summary>
            /// Fuente subrayada.
            /// </summary>
            public bool Underline
            {
                get { return underline; }
                set { underline = value; }
            }

            /// <summary>
            /// Tipo de fuente.
            /// </summary>
            public string Font
            {
                get { return font; }
                set { font = value; }
            }

            /// <summary>
            /// Tamaño de fuente.
            /// </summary>
            public int Size
            {
                get { return size; }
                set { size = value; }
            }

            /// <summary>
            /// Color de fuente.
            /// </summary>
            public Color Color
            {
                get { return color; }
                set { color = value; }
            }

            #endregion
        }
    }
}
