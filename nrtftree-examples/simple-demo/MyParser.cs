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
 * Class:		MyParser
 * Description:	Parser para la traducción de RTF a Pseudo XML. (Demo NRtfTree)
 ********************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;

namespace Net.Sgoliver.NRtfTree
{
    namespace Demo
    {
        public class MyParser : SarParser
        {
            public string doc = "";

            private bool enTexto = false;

            private bool negrita = false;
            private bool cursiva = false;
            private bool subrayado = false;

            public MyParser(string doc)
            {
                this.doc = doc;
            }

            public override void StartRtfDocument()
            {
                doc += "<doc>\r\n";
            }

            public override void EndRtfDocument()
            {
                doc += "</doc>";
            }

            public override void StartRtfGroup()
            {
                ;
            }

            public override void EndRtfGroup()
            {
                if (enTexto)
                {
                    if (negrita)
                        doc += "</b>";

                    if (cursiva)
                        doc += "</i>";

                    if (subrayado)
                        doc += "</u>";

                    negrita = false;
                    cursiva = false;
                    subrayado = false;
                }
            }

            public override void RtfControl(string key, bool hasParam, int param)
            {
                if (key == "'") //Caracter especial
                {
                    doc += translateAnsiCode(param);
                }
            }

            private string translateAnsiCode(int cod)
            {
                string res = "";

                switch (cod)
                {
                    case 193:
                        res = "Á";
                        break;
                    case 201:
                        res = "É";
                        break;
                    case 205:
                        res = "Í";
                        break;
                    case 211:
                        res = "Ó";
                        break;
                    case 218:
                        res = "Ú";
                        break;
                    case 225:
                        res = "á";
                        break;
                    case 233:
                        res = "é";
                        break;
                    case 237:
                        res = "í";
                        break;
                    case 243:
                        res = "ó";
                        break;
                    case 250:
                        res = "ú";
                        break;
                    case 241:
                        res = "ñ";
                        break;
                    case 209:
                        res = "Ñ";
                        break;
                }

                return res;
            }

            public override void RtfKeyword(string key, bool hasParam, int param)
            {
                if (key.Equals("pard"))
                    enTexto = true;

                if (enTexto)
                {
                    switch (key)
                    {
                        case "b":
                            if (!hasParam || (hasParam && param == 1))
                            {
                                doc += "<b>";
                                negrita = true;
                            }
                            else
                            {
                                doc += "</b>";
                                negrita = false;
                            }
                            break;
                        case "i":
                            if (!hasParam || (hasParam && param == 1))
                            {
                                doc += "<i>";
                                cursiva = true;
                            }
                            else
                            {
                                doc += "</i>";
                                cursiva = false;
                            }
                            break;
                        case "ul":
                            doc += "<u>";
                            subrayado = true;
                            break;
                        case "ulnone":
                            doc += "</u>";
                            subrayado = false;
                            break;
                        case "par":
                            doc += "\r\n";
                            break;
                    }
                }
            }

            public override void RtfText(string text)
            {
                if (enTexto)
                {
                    doc += text;
                }
            }
        }
    }
}
