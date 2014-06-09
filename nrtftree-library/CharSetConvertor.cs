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
 * Class:		CharSetConvertor
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
        /// A converter between charset and codepage. 
        /// Based on RTF Specification version 1.9.1.
        /// </summary>
        public static class CharSetConvertor
        {
            /// <summary>
            /// Get code page code.
            /// </summary>
            /// <param name="charSet"></param>
            /// <param name="defaultCodePage"></param>
            /// <returns></returns>
            public static int ToCodePage(int charSet, int defaultCodePage)
            {
                return Convert(CodeTableColumn.CharSet, charSet, defaultCodePage);
            }

            /// <summary>
            /// Get character set code.
            /// </summary>
            /// <param name="codePage"></param>
            /// <param name="defaultCharSet"></param>
            /// <returns></returns>
            public static int ToCharSet(int codePage, int defaultCharSet)
            {
                return Convert(CodeTableColumn.CodePage, codePage, defaultCharSet);
            }

            private static int Convert(CodeTableColumn fromType, int fromValue, int defaultValue)
            {
                int ix_from = (int)fromType;
                int ix_to = (int)(fromType == CodeTableColumn.CharSet ? CodeTableColumn.CodePage : CodeTableColumn.CharSet);
                for(int i=0; i<_codeTable.GetLength(0); ++i)
                {
                    if(fromValue == _codeTable[i, ix_from])
                        return _codeTable[i, ix_to];
                }
                return defaultValue;
            }

            private enum CodeTableColumn { CharSet=IX_CHARSET, CodePage=IX_CODEPAGE }
            private const int IX_CHARSET=0;
            private const int IX_CODEPAGE=1;
            private static int[,] _codeTable = new int[,]
            {//charset   codepage    Win/Mac name
                {0,	1252},	//ANSI
                {1,	0},	//Default
                {2,	42},	//Symbol
                {77,	10000},	//Mac Roman
                {78,	10001},	//Mac Shift Jis
                {79,	10003},	//Mac Hangul
                {80,	10008},	//Mac GB2312
                {81,	10002},	//Mac Big5
                //{82,	},	//Mac Johab (old)
                {83,	10005},	//Mac Hebrew
                {84,	10004},	//Mac Arabic
                {85,	10006},	//Mac Greek
                {86,	10081},	//Mac Turkish
                {87,	10021},	//Mac Thai
                {88,	10029},	//Mac East Europe
                {89,	10007},	//Mac Russian
                {128,	932},	//Shift JIS
                {129,	949},	//Hangul
                {130,	1361},	//Johab
                {134,	936},	//GB2312
                {136,	950},	//Big5
                {161,	1253},	//Greek
                {162,	1254},	//Turkish
                {163,	1258},	//Vietnamese
                {177,	1255},	//Hebrew
                {178,	1256},	//Arabic 
                //{179,	},	//Arabic Traditional (old)
                //{180,	},	//Arabic user (old)
                //{181,	},	//Hebrew user (old)
                {186,	1257},	//Baltic
                {204,	1251},	//Russian
                {222,	874},	//Thai
                {238,	1250},	//Eastern European
                {254,	437},	//PC 437
                {255,	850},	//OEM
            };
        }

    }
}
