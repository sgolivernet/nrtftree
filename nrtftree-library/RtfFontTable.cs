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
 * Class:		RtfFontTable
 * Description:	Tabla de Fuentes de un documento RTF.
 * ******************************************************************************/

using System.Collections.Generic;
using System.Collections;
using System;
using System.Text;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Tabla de fuentes de un documento RTF.
        /// </summary>
        public class RtfFontTable
        {
            /// <summary>
            /// Lista interna de fuentes.
            /// </summary>
            private Dictionary<int,RtfFont> fonts;

            /// <summary>
            /// Constructor de la clase RtfFontTable.
            /// </summary>
            public RtfFontTable()
            {
                fonts = new Dictionary<int,RtfFont>();
            }

            /// <summary>
            /// Inserta un nueva fuente en la tabla de fuentes.
            /// </summary>
            /// <param name="name">Nueva fuente a insertar.</param>
            public void AddFont(string name)
            {
                fonts.Add(newFontIndex(), new RtfFont(name, Encoding.Default.CodePage));
            }

            /// <summary>
            /// Inserta un nueva fuente en la tabla de fuentes.
            /// </summary>
            /// <param name="index">Indice de la fuente a insertar.</param>
            /// <param name="name">Nueva fuente a insertar.</param>
            public void AddFont(int index, string name)
            {
                fonts.Add(index, new RtfFont(name, Encoding.Default.CodePage));
            }

            /// <summary>
            /// Inserta un nueva fuente en la tabla de fuentes.
            /// </summary>
            /// <param name="index">Indice de la fuente a insertar.</param>
            /// <param name="RtfFont">Nueva fuente a insertar.</param>
            public void AddFont(int index, RtfFont font)
            {
                fonts.Add(index, font);
            }

            /// <summary>
            /// Obtiene la fuente n-ésima de la tabla de fuentes.
            /// </summary>
            /// <param name="index">Indice de la fuente a recuperar.</param>
            /// <returns>Fuente n-ésima de la tabla de fuentes.</returns>
            public RtfFont this[int index]
            {
                get
                {
                    return fonts[index];
                }
            }

            /// <summary>
            /// Número de fuentes en la tabla.
            /// </summary>
            public int Count
            {
                get 
                {
                    return fonts.Count;
                }
            }

            /// <summary>
            /// Obtiene el índice de una fuente determinado en la tabla.
            /// </summary>
            /// <param name="name">Fuente a consultar.</param>
            /// <returns>Indice de la fuente consultada.</returns>
            public int IndexOf(string name)
            {
                int intIndex = -1;
                IEnumerator fntIndex = fonts.GetEnumerator();

                fntIndex.Reset();
                while (fntIndex.MoveNext())
                {
                    KeyValuePair<int,RtfFont> item = (KeyValuePair<int, RtfFont>)fntIndex.Current;
                    if(String.Equals(name, ((RtfFont)item.Value).Name, System.StringComparison.OrdinalIgnoreCase))
                    {
                        intIndex = item.Key;
                        break;
                    }
                }

                return intIndex;
            }

            /// <summary>
            /// Obtiene un índice que no se esté usando en la tabla.
            /// </summary>
            /// <returns>Obtiene un índice que no se esté usando en la tabla.</returns>
            private int newFontIndex()
            {
                int intIndex = -1;
                IEnumerator fntIndex = fonts.GetEnumerator();

                fntIndex.Reset();
                while(fntIndex.MoveNext())
                {
                    KeyValuePair<int,RtfFont> item = (KeyValuePair<int, RtfFont>)fntIndex.Current;
                    if (item.Key > intIndex)
                        intIndex = item.Key;
                }

                return (intIndex + 1);
            }
        }
    }
}
