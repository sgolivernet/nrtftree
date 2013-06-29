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
 * Class:		RtfDocument
 * Description:	Clase para la generación de documentos RTF.
 * ******************************************************************************/

using System;
using System.Text;
using System.Drawing;
using System.Globalization;
using System.IO;
using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree
{
    namespace Util
    {
        /// <summary>
        /// Clase para la generación de documentos RTF.
        /// </summary>
        public class RtfDocument
        {
            #region Atributos privados

            /// <summary>
            /// Codificación del documento.
            /// </summary>
            private Encoding encoding;

            /// <summary>
            /// Tabla de fuentes del documento.
            /// </summary>
            private RtfFontTable fontTable;

            /// <summary>
            /// Tabla de colores del documento.
            /// </summary>
            private RtfColorTable colorTable;

            /// <summary>
            /// Grupo principal del documento.
            /// </summary>
            private RtfTreeNode mainGroup;

            /// <summary>
            /// Formato actual de caracter.
            /// </summary>
            private RtfCharFormat currentFormat;

            /// <summary>
            /// Formato actual de párrafo.
            /// </summary>
            private RtfParFormat currentParFormat;

            /// <summary>
            /// Formato del documento.
            /// </summary>
            private RtfDocumentFormat docFormat;

            #endregion

            #region Constructores

            /// <summary>
            /// Constructor de la clase RtfDocument.
            /// </summary>
            /// <param name="enc">Codificación del documento a generar.</param>
            public RtfDocument(Encoding enc)
            {
                encoding = enc;

                fontTable = new RtfFontTable();
                fontTable.AddFont("Arial");  //Default font

                colorTable = new RtfColorTable();
                colorTable.AddColor(Color.Black);  //Default color

                currentFormat = null;
                currentParFormat = new RtfParFormat();
                docFormat = new RtfDocumentFormat();

                mainGroup = new RtfTreeNode(RtfNodeType.Group);

                InitializeTree();
            }

            /// <summary>
            /// Constructor de la clase RtfDocument. Se utilizará la codificación por defecto del sistema.
            /// </summary>
            public RtfDocument() : this(Encoding.Default)
            {
            }

            #endregion

            #region Metodos Publicos

            /// <summary>
            /// Guarda el documento como fichero RTF en la ruta indicada.
            /// </summary>
            /// <param name="path">Ruta del fichero a crear.</param>
            public void Save(string path)
            {
                RtfTree tree = GetTree();
                tree.SaveRtf(path);
            }

            /// <summary>
            /// Inserta un fragmento de texto en el documento con un formato de texto determinado.
            /// </summary>
            /// <param name="text">Texto a insertar.</param>
            /// <param name="format">Formato del texto a insertar.</param>
            public void AddText(string text, RtfCharFormat format)
            {
                UpdateFontTable(format);
                UpdateColorTable(format);

                UpdateCharFormat(format);

                InsertText(text);
            }

            /// <summary>
            /// Inserta un fragmento de texto en el documento con el formato de texto actual.
            /// </summary>
            /// <param name="text">Texto a insertar.</param>
            public void AddText(string text)
            {
                InsertText(text);
            }

            /// <summary>
            /// Inserta un número determinado de saltos de línea en el documento.
            /// </summary>
            /// <param name="n">Número de saltos de línea a insertar.</param>
            public void AddNewLine(int n)
            {
                for(int i=0; i<n; i++)
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
            }

            /// <summary>
            /// Inserta un salto de línea en el documento.
            /// </summary>
            public void AddNewLine()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
            }

            /// <summary>
            /// Inicia un nuevo párrafo.
            /// </summary>
            public void AddNewParagraph()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));
            }

            /// <summary>
            /// Inserta un número determinado de saltos de párrafo en el documento.
            /// </summary>
            /// <param name="n">Número de saltos de párrafo a insertar.</param>
            public void AddNewParagraph(int n)
            {
                for (int i = 0; i < n; i++)
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));
            }

            /// <summary>
            /// Inicia un nuevo párrafo con el formato especificado.
            /// </summary>
            public void AddNewParagraph(RtfParFormat format)
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));

                UpdateParFormat(format);
            }

            /// <summary>
            /// Inserta una imagen en el documento.
            /// </summary>
            /// <param name="path">Ruta de la imagen a insertar.</param>
            /// <param name="width">Ancho deseado de la imagen en el documento.</param>
            /// <param name="height">Alto deseado de la imagen en el documento.</param>
            public void AddImage(string path, int width, int height)
            {
                FileStream fStream = null;
                BinaryReader br = null;

                try
                {
                    byte[] data = null;

                    FileInfo fInfo = new FileInfo(path);
                    long numBytes = fInfo.Length;

                    fStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                    br = new BinaryReader(fStream);

                    data = br.ReadBytes((int)numBytes);

                    StringBuilder hexdata = new StringBuilder();

                    for (int i = 0; i < data.Length; i++)
                    {
                        hexdata.Append(GetHexa(data[i]));
                    }

                    Image img = Image.FromFile(path);

                    RtfTreeNode imgGroup = new RtfTreeNode(RtfNodeType.Group);
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pict", false, 0));

                    string format = "";
                    if (path.ToLower().EndsWith("wmf"))
                        format = "emfblip";
                    else
                        format = "jpegblip";

                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, format, false, 0));
                    
                    
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "picw", true, img.Width * 20));
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pich", true, img.Height * 20));
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "picwgoal", true, width * 20));
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pichgoal", true, height * 20));
                    imgGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, hexdata.ToString(), false, 0));

                    mainGroup.AppendChild(imgGroup);
                }
                finally
                {
                    if (br != null) br.Close();
                    if (fStream != null) fStream.Close();
                }
            }

            /// <summary>
            /// Establece el formato de caracter por defecto.
            /// </summary>
            public void ResetCharFormat()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));
            }

            /// <summary>
            /// Establece el formato de párrafo por defecto.
            /// </summary>
            public void ResetParFormat()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
            }

            /// <summary>
            /// Establece el formato de caracter y párrafo por defecto.
            /// </summary>
            public void ResetFormat()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));
            }

            /// <summary>
            /// Actualiza los valores de las propiedades de formato de documento.
            /// </summary>
            /// <param name="format">Formato de documento.</param>
            public void UpdateDocFormat(RtfDocumentFormat format)
            {
                docFormat.MarginL = format.MarginL;
                docFormat.MarginR = format.MarginR;
                docFormat.MarginT = format.MarginT;
                docFormat.MarginB = format.MarginB;
            }

            /// <summary>
            /// Actualiza los valores de las propiedades de formato de texto y párrafo.
            /// </summary>
            /// <param name="format">Formato de texto a insertar.</param>
            public void UpdateCharFormat(RtfCharFormat format)
            {
                if (currentFormat != null)
                {
                    SetFormatColor(format.Color);
                    SetFormatSize(format.Size);
                    SetFormatFont(format.Font);

                    SetFormatBold(format.Bold);
                    SetFormatItalic(format.Italic);
                    SetFormatUnderline(format.Underline);
                }
                else //currentFormat == null
                {
                    int indColor = colorTable.IndexOf(format.Color);

                    if (indColor == -1)
                    {
                        colorTable.AddColor(format.Color);
                        indColor = colorTable.IndexOf(format.Color);
                    }

                    int indFont = fontTable.IndexOf(format.Font);

                    if (indFont == -1)
                    {
                        fontTable.AddFont(format.Font);
                        indFont = fontTable.IndexOf(format.Font);
                    }

                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "cf", true, indColor));
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fs", true, format.Size * 2));
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, indFont));

                    if (format.Bold)
                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "b", false, 0));

                    if (format.Italic)
                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "i", false, 0));

                    if (format.Underline)
                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ul", false, 0));

                    currentFormat = new RtfCharFormat();
                    currentFormat.Color = format.Color;
                    currentFormat.Size = format.Size;
                    currentFormat.Font = format.Font;
                    currentFormat.Bold = format.Bold;
                    currentFormat.Italic = format.Italic;
                    currentFormat.Underline = format.Underline;
                }
            }

            /// <summary>
            /// Establece el formato de párrafo pasado como parámetro.
            /// </summary>
            /// <param name="format">Formato de párrafo a utilizar.</param>
            public void UpdateParFormat(RtfParFormat format)
            {
                SetAlignment(format.Alignment);
                SetLeftIndentation(format.LeftIndentation);
                SetRightIndentation(format.RightIndentation);
            }

            /// <summary>
            /// Estable la alineación del texto dentro del párrafo.
            /// </summary>
            /// <param name="align">Tipo de alineación.</param>
            public void SetAlignment(TextAlignment align)
            {
                if (currentParFormat.Alignment != align)
                {
                    string keyword = "";

                    switch (align)
                    { 
                        case TextAlignment.Left:
                            keyword = "ql";
                            break;
                        case TextAlignment.Right:
                            keyword = "qr";
                            break;
                        case TextAlignment.Centered:
                            keyword = "qc";
                            break;
                        case TextAlignment.Justified:
                            keyword = "qj";
                            break;
                    }

                    currentParFormat.Alignment = align;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, keyword, false, 0));
                }
            }

            /// <summary>
            /// Establece la sangría izquierda del párrafo.
            /// </summary>
            /// <param name="val">Sangría izquierda en centímetros.</param>
            public void SetLeftIndentation(float val)
            {
                if (currentParFormat.LeftIndentation != val)
                {
                    currentParFormat.LeftIndentation = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "li", true, calcTwips(val)));
                }
            }

            /// <summary>
            /// Establece la sangría derecha del párrafo.
            /// </summary>
            /// <param name="val">Sangría derecha en centímetros.</param>
            public void SetRightIndentation(float val)
            {
                if (currentParFormat.RightIndentation != val)
                {
                    currentParFormat.RightIndentation = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ri", true, calcTwips(val)));
                }
            }

            /// <summary>
            /// Establece el indicativo de fuente negrita.
            /// </summary>
            /// <param name="val">Indicativo de fuente negrita.</param>
            public void SetFormatBold(bool val)
            {
                if (currentFormat.Bold != val)
                {
                    currentFormat.Bold = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "b", val ? false : true, 0));
                }
            }

            /// <summary>
            /// Establece el indicativo de fuente cursiva.
            /// </summary>
            /// <param name="val">Indicativo de fuente cursiva.</param>
            public void SetFormatItalic(bool val)
            {
                if (currentFormat.Italic != val)
                {
                    currentFormat.Italic = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "i", val ? false : true, 0));
                }
            }

            /// <summary>
            /// Establece el indicativo de fuente subrayada.
            /// </summary>
            /// <param name="val">Indicativo de fuente subrayada.</param>
            public void SetFormatUnderline(bool val)
            {
                if (currentFormat.Underline != val)
                {
                    currentFormat.Underline = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ul", val ? false : true, 0));
                }
            }

            /// <summary>
            /// Establece el color de fuente actual.
            /// </summary>
            /// <param name="val">Color de la fuente.</param>
            public void SetFormatColor(Color val)
            {
                if (currentFormat.Color.ToArgb() != val.ToArgb())
                {
                    int indColor = colorTable.IndexOf(val);

                    if (indColor == -1)
                    {
                        colorTable.AddColor(val);
                        indColor = colorTable.IndexOf(val);
                    }

                    currentFormat.Color = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "cf", true, indColor));
                }
            }

            /// <summary>
            /// Establece el tamaño de fuente actual.
            /// </summary>
            /// <param name="val">Tamaño de la fuente.</param>
            public void SetFormatSize(int val)
            {
                if (currentFormat.Size != val)
                {
                    currentFormat.Size = val;

                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fs", true, val * 2));
                }
            }

            /// <summary>
            /// Establece el tipo de letra actual.
            /// </summary>
            /// <param name="val">Tipo de letra.</param>
            public void SetFormatFont(string val)
            {
                if (currentFormat.Font != val)
                {
                    int indFont = fontTable.IndexOf(val);

                    if (indFont == -1)
                    {
                        fontTable.AddFont(val);
                        indFont = fontTable.IndexOf(val);
                    }

                    currentFormat.Font = val;
                    mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, indFont));
                }
            }

            #endregion

            #region Propiedades

            /// <summary>
            /// Obtiene el texto plano contenido en el documento RTF
            /// </summary>
            public string Text
            {
                get 
                {
                    return GetTree().Text;
                }
            }

            /// <summary>
            /// Obtiene el código RTF del documento RTF
            /// </summary>
            public string Rtf
            {
                get
                {
                    return GetTree().Rtf;
                }
            }

            /// <summary>
            /// Obtiene el árbol RTF del documento actual
            /// </summary>
            public RtfTree Tree
            {
                get
                {
                    return GetTree();
                }
            }

            #endregion

            #region Metodos Privados

            /// <summary>
            /// Obtiene el árbol RTF equivalente al documento actual.
            /// <returns>Árbol RTF equivalente al documento en el estado actual.</returns>
            /// </summary>
            private RtfTree GetTree()
            {
                RtfTree tree = new RtfTree();
                tree.RootNode.AppendChild(mainGroup.CloneNode());

                InsertFontTable(tree);
                InsertColorTable(tree);
                InsertGenerator(tree);
                InsertDocSettings(tree);

                tree.MainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "par", false, 0));

                return tree;
            }

            /// <summary>
            /// Obtiene el código hexadecimal de un entero.
            /// </summary>
            /// <param name="code">Número entero.</param>
            /// <returns>Código hexadecimal del entero pasado como parámetro.</returns>
            private string GetHexa(byte code)
            {
                string hexa = Convert.ToString(code, 16);

                if (hexa.Length == 1)
                {
                    hexa = "0" + hexa;
                }

                return hexa;
            }

            /// <summary>
            /// Inserta el código RTF de la tabla de fuentes en el documento.
            /// </summary>
            private void InsertFontTable(RtfTree tree)
            {
                RtfTreeNode ftGroup = new RtfTreeNode(RtfNodeType.Group);
                
                ftGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fonttbl", false, 0));

                for(int i=0; i<fontTable.Count; i++)
                {
                    RtfTreeNode ftFont = new RtfTreeNode(RtfNodeType.Group);
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, i));
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fnil", false, 0));
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Text, fontTable[i] + ";", false, 0));

                    ftGroup.AppendChild(ftFont);
                }

                tree.MainGroup.InsertChild(5, ftGroup);
            }

            /// <summary>
            /// Inserta el código RTF de la tabla de colores en el documento.
            /// </summary>
            private void InsertColorTable(RtfTree tree)
            {
                RtfTreeNode ctGroup = new RtfTreeNode(RtfNodeType.Group);

                ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "colortbl", false, 0));

                for (int i = 0; i < colorTable.Count; i++)
                {
                    ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "red", true, colorTable[i].R));
                    ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "green", true, colorTable[i].G));
                    ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "blue", true, colorTable[i].B));
                    ctGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, ";", false, 0));
                }

                tree.MainGroup.InsertChild(6, ctGroup);
            }

            /// <summary>
            /// Inserta el código RTF de la aplicación generadora del documento.
            /// </summary>
            private void InsertGenerator(RtfTree tree)
            {
                RtfTreeNode genGroup = new RtfTreeNode(RtfNodeType.Group);

                genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Control, "*", false, 0));
                genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "generator", false, 0));
                genGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, "NRtfTree Library 0.3.0;", false, 0));

                tree.MainGroup.InsertChild(7, genGroup);
            }

            /// <summary>
            /// Inserta todos los nodos de texto y control necesarios para representar un texto determinado.
            /// </summary>
            /// <param name="text">Texto a insertar.</param>
            private void InsertText(string text)
            {
                int i = 0;
                int code = 0;

                while(i < text.Length)
                {
                    code = Char.ConvertToUtf32(text, i);

                    if(code >= 32 && code < 128)
                    {
                        StringBuilder s = new StringBuilder("");

                        while (i < text.Length && code >= 32 && code < 128)
                        {
                            s.Append(text[i]);

                            i++;

                            if(i < text.Length)
                                code = Char.ConvertToUtf32(text, i);
                        }

                        mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, s.ToString(), false, 0));
                    }
                    else
                    {
                        if (text[i] == '\t')
                        {
                            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "tab", false, 0));
                        }
                        else if (text[i] == '\n')
                        {
                            mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "line", false, 0));
                        }
                        else
                        {
                            if (code <= 255)
                            {
                                byte[] bytes = encoding.GetBytes(new char[] { text[i] });

                                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Control, "'", true, bytes[0]));
                            }
                            else
                            {
                                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "u", true, code));
                                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Text, "?", false, 0));
                            }
                        }

                        i++;
                    }
                }
            }

            /// <summary>
            /// Actualiza la tabla de fuentes con una nueva fuente si es necesario.
            /// </summary>
            /// <param name="format"></param>
            private void UpdateFontTable(RtfCharFormat format)
            {
                if (fontTable.IndexOf(format.Font) == -1)
                {
                    fontTable.AddFont(format.Font);
                }
            }

            /// <summary>
            /// Actualiza la tabla de colores con un nuevo color si es necesario.
            /// </summary>
            /// <param name="format"></param>
            private void UpdateColorTable(RtfCharFormat format)
            {
                if (colorTable.IndexOf(format.Color) == -1)
                {
                    colorTable.AddColor(format.Color);
                }
            }

            /// <summary>
            /// Inicializa el arbol RTF con todas las claves de la cabecera del documento.
            /// </summary>
            private void InitializeTree()
            {
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "rtf", true, 1));
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ansi", false, 0));
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "ansicpg", true, encoding.CodePage));
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "deff", true, 0));
                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "deflang", true, CultureInfo.CurrentCulture.LCID));

                mainGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
            }

            /// <summary>
            /// Inserta las propiedades de formato del documento
            /// </summary>
            private void InsertDocSettings(RtfTree tree)
            {
                int indInicioTexto = tree.MainGroup.ChildNodes.IndexOf("pard");

                //Generic Properties
                
                tree.MainGroup.InsertChild(indInicioTexto, new RtfTreeNode(RtfNodeType.Keyword, "viewkind", true, 4));
                tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "uc", true, 1));

                //RtfDocumentFormat Properties

                tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margl", true, calcTwips(docFormat.MarginL)));
                tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margr", true, calcTwips(docFormat.MarginR)));
                tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margt", true, calcTwips(docFormat.MarginT)));
                tree.MainGroup.InsertChild(indInicioTexto++, new RtfTreeNode(RtfNodeType.Keyword, "margb", true, calcTwips(docFormat.MarginB)));
            }

            /// <summary>
            /// Convierte entre centímetros y twips.
            /// </summary>
            /// <param name="centimeters">Valor en centímetros.</param>
            /// <returns>Valor en twips.</returns>
            private int calcTwips(float centimeters)
            {
                //1 inches = 2.54 centimeters
                //1 inches = 1440 twips
                //20 twip = 1 pixel

                //X centimetros --> (X*1440)/2.54 twips

                return (int)((centimeters * 1440F) / 2.54F);
            }

            #endregion
        }
    }
}
