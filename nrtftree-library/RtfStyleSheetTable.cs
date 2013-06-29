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
 * Class:		RtfStyleSheetTable
 * Description:	Representa la tabla de hojas de estilo de un documento RTF.
 * ******************************************************************************/

using System.Collections.Generic;
using System.Collections;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Representa la tabla de estilos de un documento RTF.
        /// </summary>
        public class RtfStyleSheetTable
        {
            private Dictionary<int, RtfStyleSheet> stylesheets = null;

            /// <summary>
            /// Constructor de la tabla de estilos.
            /// </summary>
            public RtfStyleSheetTable()
            {
                stylesheets = new Dictionary<int, RtfStyleSheet>();
            }

            /// <summary>
            /// Añade un nuevo estilo a la tabla de estilos. El estilo se añadirá con un nuevo índice no existente en la tabla.
            /// </summary>
            /// <param name="ss">Nuevo estilo a añadir a la tabla.</param>
            public void AddStyleSheet(RtfStyleSheet ss)
            {
                ss.Index = newStyleSheetIndex();

                stylesheets.Add(ss.Index, ss);
            }

            /// <summary>
            /// Añade un nuevo estilo a la tabla de estilos. El estilo se añadirá con el índice de estilo pasado como parámetro.
            /// </summary>
            /// <param name="index">Indice del estilo a añadir a la tabla.</param>
            /// <param name="ss">Nuevo estilo a añadir a la tabla.</param>
            public void AddStyleSheet(int index, RtfStyleSheet ss)
            {
                ss.Index = index;

                stylesheets.Add(index, ss);
            }

            /// <summary>
            /// Elimina un estilo de la tabla de estilos por índice.
            /// </summary>
            /// <param name="index"></param>
            public void RemoveStyleSheet(int index)
            {
                stylesheets.Remove(index);
            }

            /// <summary>
            /// Elimina de la tabla de estilos el estilo pasado como parámetro.
            /// </summary>
            /// <param name="ss"></param>
            public void RemoveStyleSheet(RtfStyleSheet ss)
            {
                stylesheets.Remove(ss.Index);
            }

            /// <summary>
            /// Recupera un estilo de la tabla de estilos por índice.
            /// </summary>
            /// <param name="index">Indice del estilo a recuperar.</param>
            /// <returns>Estilo cuyo índice es el pasado como parámetro.</returns>
            public RtfStyleSheet GetStyleSheet(int index)
            {
                return stylesheets[index];
            }

            /// <summary>
            /// Recupera un estilo de la tabla de estilos por índice.
            /// </summary>
            /// <param name="index">Indice del estilo a recuperar.</param>
            /// <returns>Estilo cuyo índice es el pasado como parámetro.</returns>
            public RtfStyleSheet this[int index]
            {
                get
                {
                    return stylesheets[index];
                }
            }

            /// <summary>
            /// Número de estilos contenidos en la tabla de estilos.
            /// </summary>
            public int Count
            {
                get
                {
                    return stylesheets.Count;
                }
            }

            /// <summary>
            /// Índice del estilo cuyo nombre es el pasado como parámetro.
            /// </summary>
            /// <param name="name">Nombre del estilo buscado.</param>
            /// <returns>Estilo cuyo nombre es el pasado como parámetro.</returns>
            public int IndexOf(string name)
            {
                int intIndex = -1;
                IEnumerator fntIndex = stylesheets.GetEnumerator();

                fntIndex.Reset();
                while (fntIndex.MoveNext())
                {
                    if (((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Value.Name.Equals(name))
                    {
                        intIndex = (int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key;
                        break;
                    }
                }

                return intIndex;
            }

            /// <summary>
            /// Calcula un nuevo índice para insertar un estilo en la tabla.
            /// </summary>
            /// <returns>Índice del próximo estilo a insertar.</returns>
            private int newStyleSheetIndex()
            {
                int intIndex = -1;
                IEnumerator fntIndex = stylesheets.GetEnumerator();

                fntIndex.Reset();
                while (fntIndex.MoveNext())
                {
                    if ((int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key > intIndex)
                        intIndex = (int)((KeyValuePair<int, RtfStyleSheet>)fntIndex.Current).Key;
                }

                return (intIndex + 1);
            }
        }
    }
}
