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
 * Class:		Rtf2Html
 * Description:	Traducción de documentos RTF a formato HTML.
 * Notes:       Contribución de Francisco Javier Marín (http://www.xuletas.es/).
 ********************************************************************************/

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;

namespace Net.Sgoliver.NRtfTree
{
    namespace Demo
    {
        /// <summary>
        /// Conversor de documentos RTF a formato HTML.
        /// </summary>
        public class Rtf2Html
        {
            #region Atributos Privados

            /// <summary>
            /// StringBuilder empleado para escribir el código HTML resultado
            /// </summary>
            private StringBuilder _builder;

            /// <summary>
            /// Formato de texto que se está leyendo
            /// </summary>
            private Format _currentFormat;

            private Encoding _currentEncoding = Encoding.Default;

            /// <summary>
            /// Formato de texto ya escrito en el código HTML
            /// </summary>
            private Format _htmlFormat;

            /// <summary>
            /// Tabla de colores RTF
            /// </summary>
            private RtfColorTable _colorTable;

            /// <summary>
            /// Tabla de fuentes RTF
            /// </summary>
            private RtfFontTable _fontTable;

            //Propiedades

            private bool _autoParagraph;
            private bool _ignoreFontNames;
            private bool _escapeHtmlEntities;
            private string _defaultFontName;
            private int _defaultFontSize;
            private bool _incrustImages;
            private string _imagePath;
            private ImageFormat _imageFormat;

            #endregion

            #region Constructores

            public Rtf2Html()
            {
                AutoParagraph = true;
                IgnoreFontNames = false;
                DefaultFontSize = 10;
                EscapeHtmlEntities = true;
                DefaultFontName = "Times New Roman";

                ImageFormat = ImageFormat.Jpeg;
                IncrustImages = true;
                ImagePath = "";
            }

            #endregion

            #region Propiedades

            /// <summary>
            /// Obtiene o establece un valor que indica si se crearán parrafos automáticamente
            /// </summary>
            public bool AutoParagraph
            {
                get
                {
                    return _autoParagraph;
                }
                set
                {
                    _autoParagraph = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica si se ignorarán los nombres de las fuentes.
            /// Establecer este valor como false generará un archivo HTML sin propiedad CSS font-face
            /// </summary>
            public bool IgnoreFontNames
            {
                get
                {
                    return _ignoreFontNames;
                }
                set
                {
                    _ignoreFontNames = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica si se reemplazarán las entidades HTML
            /// del texto
            /// </summary>
            public bool EscapeHtmlEntities
            {
                get
                {
                    return _escapeHtmlEntities;
                }
                set
                {
                    _escapeHtmlEntities = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica el nombre de la fuente por defecto.
            /// Los grupos de textos que usen esta fuente no incluirán la propiedad font-face
            /// </summary>
            public string DefaultFontName
            {
                get
                {
                    return _defaultFontName;
                }
                set
                {
                    _defaultFontName = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica el tamaño de fuente que se ignorará.
            /// Los grupos de texto que tengan ese tamaño no incluirán la propiedad CSS font-size
            /// </summary>
            public int DefaultFontSize
            {
                get
                {
                    return _defaultFontSize;
                }
                set
                {
                    _defaultFontSize = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica si se incrustarán las imágenes en base64
            /// dentro del código HTML  (true), o se guardarán en un archivo (false)
            /// </summary>
            /// <see cref="http://www.sweeting.org/mark/blog/2005/07/12/base64-encoded-images-embedded-in-html"/>
            public bool IncrustImages
            {
                get
                {
                    return _incrustImages;
                }
                set
                {
                    _incrustImages = value;
                }
            }

            /// <summary>
            /// Obtiene o establece la ruta donde se guardarán las imágenes contenidas en el
            /// código RTF original. Sólo se usarará si el valor de IncrustImages es false
            /// </summary>
            public string ImagePath
            {
                get
                {
                    return _imagePath;
                }
                set
                {
                    _imagePath = value;
                }
            }

            /// <summary>
            /// Obtiene o establece un valor que indica el formato en el que se guardarán las imágenes
            /// incrustadas en el código RTF convertido
            /// </summary>
            public ImageFormat ImageFormat
            {
                get
                {
                    return _imageFormat;
                }
                set
                {
                    _imageFormat = value;
                }
            }

            #endregion

            #region Métodos Públicos

            /// <summary>
            /// Convierte una cadena de código RTF a formato HTML
            /// </summary>
            public static string ConvertCode(string rtf)
            {
                Rtf2Html rtfToHtml = new Rtf2Html();
                return rtfToHtml.Convert(rtf);
            }

            /// <summary>
            /// Convierte una cadena de código RTF a formato HTML
            /// </summary>
            public string Convert(string rtf)
            {
                //Generar arbol DOM
                RtfTree rtfTree = new RtfTree();
                rtfTree.LoadRtfText(rtf);

                //Inicializar variables empleadas
                _builder = new StringBuilder();
                _htmlFormat = new Format();
                _currentFormat = new Format();
                _fontTable = rtfTree.GetFontTable();
                _colorTable = rtfTree.GetColorTable();

                //Buscar el inicio del contenido visible del documento
                int inicio;
                for (inicio = 0; inicio < rtfTree.RootNode.FirstChild.ChildNodes.Count; inicio++)
                {
                    if (rtfTree.RootNode.FirstChild.ChildNodes[inicio].NodeKey == "pard")
                        break;
                }

                //Procesar todos los nodos visibles
                ProcessChildNodes(rtfTree.RootNode.FirstChild.ChildNodes, inicio);

                //Cerrar etiquetas pendientes
                _currentFormat.Reset();
                WriteText(string.Empty);

                //Arreglar HTML

                //Arreglar listas
                Regex repairList = new Regex("<span [^>]*>·</span><span style=\"([^\"]*)\">(.*?)<br\\s+/><" + "/span>",
                                             RegexOptions.IgnoreCase | RegexOptions.Singleline |
                                             RegexOptions.CultureInvariant);

                foreach (Match match in repairList.Matches(_builder.ToString()))
                {
                    _builder.Replace(match.Value, string.Format("<li style=\"{0}\">{1}</li>", match.Groups[1].Value, match.Groups[2].Value));
                }

                Regex repairUl = new Regex("(?<!</li>)<li", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                foreach (Match match in repairUl.Matches(_builder.ToString()))
                {
                    _builder.Insert(match.Index, "<ul>");
                }

                repairUl = new Regex("/li>(?!<li)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                foreach (Match match in repairUl.Matches(_builder.ToString()))
                {
                    _builder.Insert(match.Index + match.Length, "</ul>");
                }

                //Generar párrafos (cada 2 <br /><br /> se cambiará por un <p>)
                if (AutoParagraph)
                {
                    string[] partes = _builder.ToString().Split(new[] { "<br /><br />" }, StringSplitOptions.RemoveEmptyEntries);
                    _builder = new StringBuilder(_builder.Length + 7 * partes.Length);

                    foreach (string parte in partes)
                    {
                        _builder.Append("<p>");
                        _builder.Append(parte);
                        _builder.Append("</p>");
                    }
                }

                return EscapeHtmlEntities ? HtmlEntities.Encode(_builder.ToString()) : _builder.ToString();
            }

            #endregion

            #region Métodos Privados

            /// <summary>
            /// Función encargada de procesar los nodos hijo de un nodo padre RTF un nodo del documento RTF
            /// y generar las etiquetas HTML necesarias
            /// </summary>
            private void ProcessChildNodes(RtfNodeCollection nodos, int inicio)
            {
                for (int i = inicio; i < nodos.Count; i++)
                {
                    RtfTreeNode nodo = nodos[i];

                    switch (nodo.NodeType)
                    {
                        case RtfNodeType.Control:

                            if (nodo.NodeKey == "'") //Símbolos especiales, como tildes y "ñ"
                            {
                                WriteText(_currentEncoding.GetString(new[] { (byte)nodo.Parameter }));
                            }
                            break;

                        case RtfNodeType.Keyword:

                            switch (nodo.NodeKey)
                            {
                                case "pard": //Reinicio de formato
                                    _currentFormat.Reset();
                                    break;

                                case "f": //Tipo de fuente                                
                                    if (nodo.Parameter < _fontTable.Count)
                                    {
                                        _currentFormat.FontName = _fontTable[nodo.Parameter].Name;
                                        _currentEncoding = Encoding.GetEncoding(_fontTable[nodo.Parameter].CodePage);
                                    }
                                    break;

                                case "cf": //Color de fuente
                                    if (nodo.Parameter < _colorTable.Count)
                                        _currentFormat.ForeColor = _colorTable[nodo.Parameter];
                                    break;

                                case "highlight": //Color de fondo
                                    if (nodo.Parameter < _colorTable.Count)
                                        _currentFormat.BackColor = _colorTable[nodo.Parameter];
                                    break;

                                case "fs": //Tamaño de fuente
                                    _currentFormat.FontSize = nodo.Parameter;
                                    break;

                                case "b": //Negrita
                                    _currentFormat.Bold = !nodo.HasParameter || nodo.Parameter == 1;
                                    break;

                                case "i": //Cursiva
                                    _currentFormat.Italic = !nodo.HasParameter || nodo.Parameter == 1;
                                    break;

                                case "ul": //Subrayado ON
                                    _currentFormat.Underline = true;
                                    break;

                                case "ulnone": //Subrayado OFF
                                    _currentFormat.Underline = false;
                                    break;

                                case "super": //Superscript
                                    _currentFormat.Superscript = true;
                                    _currentFormat.Subscript = false;
                                    break;

                                case "sub": //Subindice
                                    _currentFormat.Subscript = true;
                                    _currentFormat.Superscript = false;
                                    break;

                                case "nosupersub":
                                    _currentFormat.Superscript = _currentFormat.Subscript = false;
                                    break;

                                case "qc": //Alineacion centrada
                                    _currentFormat.Alignment = HorizontalAlignment.Center;
                                    break;

                                case "qr": //Alineacion derecha
                                    _currentFormat.Alignment = HorizontalAlignment.Right;
                                    break;

                                case "li": //tabulacion
                                    _currentFormat.Margin = nodo.Parameter;
                                    break;

                                case "line":
                                case "par": //Nueva línea
                                    _builder.Append("<br />");
                                    break;
                            }

                            break;

                        case RtfNodeType.Group:

                            //Procesar nodos hijo, si los tiene
                            if (nodo.HasChildNodes())
                            {
                                if (nodo["pict"] != null) //El grupo es una imagen
                                {
                                    ImageNode imageNode = new ImageNode(nodo);
                                    WriteImage(imageNode);
                                }
                                else
                                {
                                    ProcessChildNodes(nodo.ChildNodes, 0);
                                }
                            }
                            break;

                        case RtfNodeType.Text:

                            WriteText(nodo.NodeKey);
                            break;

                        default:

                            throw new ArgumentOutOfRangeException();
                    }
                }
            }

            /// <summary>
            /// Función encargada de añadir texto con el formato actual al código html resultado
            /// </summary>
            private void WriteText(string text)
            {
                if (_builder.Length > 0)
                {
                    //Cerrar etiquetas
                    if (_currentFormat.Bold == false && _htmlFormat.Bold)
                    {
                        _builder.Append("</strong>");
                        _htmlFormat.Bold = false;
                    }
                    if (_currentFormat.Italic == false && _htmlFormat.Italic)
                    {
                        _builder.Append("</em>");
                        _htmlFormat.Italic = false;
                    }
                    if (_currentFormat.Underline == false && _htmlFormat.Underline)
                    {
                        _builder.Append("</u>");
                        _htmlFormat.Underline = false;
                    }
                    if (_currentFormat.Subscript == false && _htmlFormat.Subscript)
                    {
                        _builder.Append("</sub>");
                        _htmlFormat.Subscript = false;
                    }
                    if (_currentFormat.Superscript == false && _htmlFormat.Superscript)
                    {
                        _builder.Append("</sup>");
                        _htmlFormat.Superscript = false;
                    }
                    if (_currentFormat.CompareFontFormat(_htmlFormat) == false) //El formato de fuente ha cambiado
                    {
                        _builder.Append("</span>");

                        //Reiniciar propiedades
                        _htmlFormat.Reset();
                    }
                }

                //Abrir etiquetas necesarias para representar el formato actual
                if (_currentFormat.CompareFontFormat(_htmlFormat) == false) //El formato de fuente ha cambiado
                {
                    string estilo = string.Empty;

                    if (!IgnoreFontNames && !string.IsNullOrEmpty(_currentFormat.FontName) &&
                        string.Compare(_currentFormat.FontName, DefaultFontName, true) != 0)
                        estilo += string.Format("font-family:{0};", _currentFormat.FontName);
                    if (_currentFormat.FontSize > 0 && _currentFormat.FontSize / 2 != DefaultFontSize)
                        estilo += string.Format("font-size:{0}pt;", _currentFormat.FontSize / 2);
                    if (_currentFormat.Margin != _htmlFormat.Margin)
                        estilo += string.Format("margin-left:{0}px;", _currentFormat.Margin / 15);
                    if (_currentFormat.Alignment != _htmlFormat.Alignment)
                        estilo += string.Format("text-align:{0};", _currentFormat.Alignment.ToString().ToLower());
                    if (CompareColor(_currentFormat.ForeColor, _htmlFormat.ForeColor) == false)
                        estilo += string.Format("color:{0};", ColorTranslator.ToHtml(_currentFormat.ForeColor));
                    if (CompareColor(_currentFormat.BackColor, _htmlFormat.BackColor) == false)
                        estilo += string.Format("background-color:{0};", ColorTranslator.ToHtml(_currentFormat.BackColor));

                    _htmlFormat.FontName = _currentFormat.FontName;
                    _htmlFormat.FontSize = _currentFormat.FontSize;
                    _htmlFormat.ForeColor = _currentFormat.ForeColor;
                    _htmlFormat.BackColor = _currentFormat.BackColor;
                    _htmlFormat.Margin = _currentFormat.Margin;
                    _htmlFormat.Alignment = _currentFormat.Alignment;

                    if (!string.IsNullOrEmpty(estilo))
                        _builder.AppendFormat("<span style=\"{0}\">", estilo);
                }
                if (_currentFormat.Superscript && _htmlFormat.Superscript == false)
                {
                    _builder.Append("<sup>");
                    _htmlFormat.Superscript = true;
                }
                if (_currentFormat.Subscript && _htmlFormat.Subscript == false)
                {
                    _builder.Append("<sub>");
                    _htmlFormat.Subscript = true;
                }
                if (_currentFormat.Underline && _htmlFormat.Underline == false)
                {
                    _builder.Append("<u>");
                    _htmlFormat.Underline = true;
                }
                if (_currentFormat.Italic && _htmlFormat.Italic == false)
                {
                    _builder.Append("<em>");
                    _htmlFormat.Italic = true;
                }
                if (_currentFormat.Bold && _htmlFormat.Bold == false)
                {
                    _builder.Append("<strong>");
                    _htmlFormat.Bold = true;
                }

                _builder.Append(text.Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;"));
            }

            /// <summary>
            /// Función encargada de añadir una imagen al código HTML resultado
            /// </summary>
            /// <param name="imageNode"></param>
            private void WriteImage(ImageNode imageNode)
            {
                Bitmap b = imageNode.Bitmap;
                Size imageSize;

                if (imageNode.DesiredHeight > 0 && imageNode.DesiredWidth > 0)//Aproximar twips a px dividiendo entre 15
                    imageSize = new Size(imageNode.DesiredWidth / 15, imageNode.DesiredHeight / 15);
                else
                    imageSize = b.Size;

                //Buscar codec
                ImageCodecInfo imageCodec = null;
                foreach (ImageCodecInfo codec in ImageCodecInfo.GetImageDecoders())
                {
                    if (imageCodec == null || codec.FormatID == ImageFormat.Guid)
                        imageCodec = codec;
                }

                //Reducir imagen si el tamaño final es menor que el actual
                if (b.Size.Height > imageSize.Height || b.Size.Width > imageSize.Width)
                {
                    Bitmap bmpResized = new Bitmap(imageSize.Width, imageSize.Height);
                    Graphics g = Graphics.FromImage(bmpResized);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(b, 0, 0, imageSize.Width, imageSize.Height);
                    b.Dispose();
                    b = bmpResized;
                }

                //Generar imagen en el formato adecuado
                string imageSrc;
                if (IncrustImages)
                {
                    using (var buffer = new MemoryStream())
                    {
                        b.Save(buffer, ImageFormat);

                        imageSrc = string.Format("data:{0};base64,{1}", imageCodec.MimeType,
                                                 System.Convert.ToBase64String(buffer.GetBuffer(), 0, (int)buffer.Length));
                    }
                }
                else
                {
                    int index = 1;
                    string ext = imageCodec.FilenameExtension.Split(';')[0].Substring(2).ToLower();
                    string path;

                    do
                    {
                        path = Path.Combine(ImagePath, string.Format("Imagen{0}.{1}", index, ext));
                        index++;
                    } while (File.Exists(path));

                    b.Save(path, ImageFormat);

                    //imageSrc = "file://" + path;
                    imageSrc = path;
                }

                //Añadir imagen al resultado HTML
                _builder.Append(string.Format("<img src=\"{0}\" style=\"width:{1}px;height:{2}px;\" />"
                                              , imageSrc, imageSize.Width, imageSize.Height));

                //Limpiar memoria
                b.Dispose();
            }

            /// <summary>
            /// Compara dos colores sin tener en cuenta el canal alfa
            /// </summary>
            private static bool CompareColor(Color a, Color b)
            {
                return a.R == b.R && a.G == b.G && a.B == b.B;
            }

            #endregion

            #region Nested type: Format

            /// <summary>
            /// Representa el formato que puede tomar un conjunto de texto
            /// </summary>
            private class Format
            {
                public bool Italic;
                public bool Bold;
                public bool Subscript;
                public bool Underline;
                public bool Superscript;

                public string FontName;
                public int FontSize;
                public Color ForeColor;
                public Color BackColor;
                public int Margin;
                public HorizontalAlignment Alignment;

                public Format()
                {
                    Reset();
                }

                /// <summary>
                /// Compara las propiedades FontName, FontSize, Margin, ForeColor, BackColor y Alignment con otro
                /// objeto de esta clase
                /// </summary>
                public bool CompareFontFormat(Format format)
                {
                    return string.Compare(FontName, format.FontName, true) == 0 &&
                           FontSize == format.FontSize &&
                           ForeColor == format.ForeColor &&
                           BackColor == format.BackColor &&
                           Margin == format.Margin &&
                           Alignment == format.Alignment;
                }

                public void Reset()
                {
                    FontName = string.Empty;
                    FontSize = 0;
                    ForeColor = Color.Black;
                    BackColor = Color.White;
                    Margin = 0;
                    Alignment = HorizontalAlignment.Left;
                }
            }

            private enum HorizontalAlignment
            {
                Left,
                Right,
                Center
            }

            #endregion
        }
    }
}