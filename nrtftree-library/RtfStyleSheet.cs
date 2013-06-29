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
 * Class:		RtfStyleSheet
 * Description:	Representa una hoja de estilo de un documento RTF.
 * ******************************************************************************/

using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Representa una hoja de estilo de un documento RTF
        /// </summary>
        public class RtfStyleSheet
        {
            #region Atributos Privados

            private int index = 0;
            private string name = "";
            private RtfStyleSheetType type = RtfStyleSheetType.Paragraph;
            private bool additive = false;
            private int basedOn = -1;
            private int next = -1;
            private bool autoUpdate = false;
            private bool hidden = false;
            private int link = -1;
            private bool locked = false;
            private bool personal = false;
            private bool compose = false;
            private bool reply = false;
            private int styrsid = -1;
            private bool semiHidden = false;
            private RtfNodeCollection keyCode = null;
            private RtfNodeCollection formatting = null;

            #endregion

            #region Propiedades Públicas

            /// <summary>
            /// Índice de la hoja de estilo
            /// </summary>
            public int Index
            {
                get { return index; }
                set { index = value; }
            }

            /// <summary>
            /// Nombre de la hoja de estilo.
            /// </summary>
            public string Name
            {
                get { return name; }
                set { name = value; }
            }

            /// <summary>
            /// Tipo de la hoja de estilo.
            /// </summary>
            public RtfStyleSheetType Type
            {
                get { return type; }
                set { type = value; }
            }

            /// <summary>
            /// En las hojas de estilo de tipo Caracter indica si el estilo indicado se va a sumar al estilo de párrafo actual, en vez de sobreescribir completamente el estilo actual.
            /// </summary>
            public bool Additive
            {
                get { return additive; }
                set { additive = value; }
            }

            /// <summary>
            /// Indica el estilo en el que está basado el estilo actual.
            /// </summary>
            public int BasedOn
            {
                get { return basedOn; }
                set { basedOn = value; }
            }

            /// <summary>
            /// Indica el estilo que se usará en el siguiente párrafo.
            /// </summary>
            public int Next
            {
                get { return next; }
                set { next = value; }
            }

            /// <summary>
            /// Indican si los estilos se actualizarán automáticamente.
            /// </summary>
            public bool AutoUpdate
            {
                get { return autoUpdate; }
                set { autoUpdate = value; }
            }

            /// <summary>
            /// Establece que el estilo actual no aparecerá en las listas desplegables de estilos.
            /// </summary>
            public bool Hidden
            {
                get { return hidden; }
                set { hidden = value; }
            }

            /// <summary>
            /// Indica el estilo al que está enlazado el actual.
            /// </summary>
            public int Link
            {
                get { return link; }
                set { link = value; }
            }

            /// <summary>
            /// Indica que el estilo actual está bloqueado.
            /// </summary>
            public bool Locked
            {
                get { return locked; }
                set { locked = value; }
            }

            /// <summary>
            /// Indica que el estilo actual es el estilo de e-mail personal.
            /// </summary>
            public bool Personal
            {
                get { return personal; }
                set { personal = value; }
            }

            /// <summary>
            /// Indica que el estilo actual es el estilo de composición de e-mail.
            /// </summary>
            public bool Compose
            {
                get { return compose; }
                set { compose = value; }
            }

            /// <summary>
            /// Indica que el estilo actual es el estilo de respuesta de e-mail.
            /// </summary>
            public bool Reply
            {
                get { return reply; }
                set { reply = value; }
            }

            /// <summary>
            /// Tied to the rsid table, N is the rsid of the author who implemented the style.
            /// </summary>
            public int Styrsid
            {
                get { return styrsid; }
                set { styrsid = value; }
            }

            /// <summary>
            /// Indica que el estilo no aparecerá en los menús desplegables.
            /// </summary>
            public bool SemiHidden
            {
                get { return semiHidden; }
                set { semiHidden = value; }
            }

            /// <summary>
            /// Indica la tecla rápida para establecer el estilo actual.
            /// </summary>
            public RtfNodeCollection KeyCode
            {
                get { return keyCode; }
                set { keyCode = value; }
            }

            /// <summary>
            /// Opciones de formato contenidas en el estilo actual.
            /// </summary>
            public RtfNodeCollection Formatting
            {
                get { return formatting; }
                set { formatting = value; }
            }

            #endregion

            #region Constructores

            /// <summary>
            /// COnstructor de la clase RtfStyleSheet.
            /// </summary>
            public RtfStyleSheet()
            {
                keyCode = null;
                formatting = new RtfNodeCollection();
            }

            #endregion
        }
    }
}
