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
 * Version:     v0.3
 * Date:		20/09/2012
 * Copyright:   2006-2012 Salvador Gomez
 * Home Page:	http://www.sgoliver.net
 * GitHub:	    https://github.com/sgolivernet/nrtftree
 * Class:		RtfPullParser
 * Description:	Pull parser para documentos RTF.
 * ******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Net.Sgoliver.NRtfTree
{
    namespace Core
    {
        public class RtfPullParser
        {
            #region Constantes

            public const int START_DOCUMENT = 0;
            public const int END_DOCUMENT = 1;
            public const int KEYWORD = 2;
            public const int CONTROL = 3;
            public const int START_GROUP = 4;
            public const int END_GROUP = 5;
            public const int TEXT = 6;

            #endregion

            #region Atributos

            private TextReader rtf;		//Fichero/Cadena de entrada RTF
            private RtfLex lex;		    //Analizador léxico para RTF
            private RtfToken tok;		//Token actual
            private int currentEvent;   //Evento actual

            #endregion

            #region Construtores

            public RtfPullParser(string path)
            {
                //Se abre el fichero de entrada
                rtf = new StreamReader(path);

                //Se crea el analizador léxico para RTF
                lex = new RtfLex(rtf);

                currentEvent = START_DOCUMENT;
            }

            #endregion

            #region Métodos Públicos

            public int GetEventType()
            {
                return currentEvent;
            }

            public int Next()
            {
                tok = lex.NextToken();

                switch (tok.Type)
                {
                    case RtfTokenType.GroupStart:
                        currentEvent = START_GROUP;
                        break;
                    case RtfTokenType.GroupEnd:
                        currentEvent = END_GROUP;
                        break;
                    case RtfTokenType.Keyword:
                        currentEvent = KEYWORD;
                        break;
                    case RtfTokenType.Control:
                        currentEvent = CONTROL;
                        break;
                    case RtfTokenType.Text:
                        currentEvent = TEXT;
                        break;
                    case RtfTokenType.Eof:
                        currentEvent = END_DOCUMENT;
                        break;
                }

                return currentEvent;
            }

            public string GetName()
            {
                return tok.Key;
            }

            public int GetParam()
            {
                return tok.Parameter;
            }

            public bool HasParam()
            {
                return tok.HasParameter;
            }

            public string GetText()
            {
                return tok.Key;
            }

            #endregion
        }
    }
}
