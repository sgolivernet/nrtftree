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
 * Class:		RtfMerger
 * Description:	Clase para combinar varios documentos RTF.
 * Notes:       Originally contributed by Fabio Borghi.
 * ******************************************************************************/

using System.Collections.Generic;
using Net.Sgoliver.NRtfTree.Util;
using System.Drawing;

namespace Net.Sgoliver.NRtfTree
{
    namespace Core
    {
        /// <summary>
        /// Clase para combinar varios documentos RTF.
        /// </summary>
        public class RtfMerger
        {
            #region Atributos privados

            private RtfTree baseRtfDoc = null;
            private bool removeLastPar;

            private Dictionary<string, RtfTree> placeHolder = null;

            #endregion

            #region Constructores

            /// <summary>
            /// Constructor de la clase RtfMerger. 
            /// </summary>
            /// <param name="templatePath">Ruta del documento plantilla.</param>
            public RtfMerger(string templatePath)
            {
                //Se carga el documento origen
                baseRtfDoc = new RtfTree();
                baseRtfDoc.LoadRtfFile(templatePath);

                //Se crea la lista de parámetros de sustitución (placeholders)
                placeHolder = new Dictionary<string, RtfTree>();
            }

            /// <summary>
            /// Constructor de la clase RtfMerger. 
            /// </summary>
            /// <param name="templateTree">Ruta del documento plantilla.</param>
            public RtfMerger(RtfTree templateTree)
            {
                //Se carga el documento origen
                baseRtfDoc = templateTree;

                //Se crea la lista de parámetros de sustitución (placeholders)
                placeHolder = new Dictionary<string, RtfTree>();
            }

            /// <summary>
            /// Constructor de la clase RtfMerger. 
            /// </summary>
            public RtfMerger()
            {
                //Se crea la lista de parámetros de sustitución (placeholders)
                placeHolder = new Dictionary<string, RtfTree>();
            }

            #endregion

            #region Métodos Públicos

            /// <summary>
            /// Asocia un nuevo parámetro de sustitución (placeholder) con la ruta del documento a insertar.
            /// </summary>
            /// <param name="ph">Nombre del placeholder.</param>
            /// <param name="path">Ruta del documento a insertar.</param>
            public void AddPlaceHolder(string ph, string path)
            {
                RtfTree tree = new RtfTree();

                int res = tree.LoadRtfFile(path);

                if (res == 0)
                {
                    placeHolder.Add(ph, tree);
                }
            }

            /// <summary>
            /// Asocia un nuevo parámetro de sustitución (placeholder) con la ruta del documento a insertar.
            /// </summary>
            /// <param name="ph">Nombre del placeholder.</param>
            /// <param name="docTree">Árbol RTF del documento a insertar.</param>
            public void AddPlaceHolder(string ph, RtfTree docTree)
            {
                placeHolder.Add(ph, docTree);
            }

            /// <summary>
            /// Desasocia un parámetro de sustitución (placeholder) con la ruta del documento a insertar.
            /// </summary>
            /// <param name="ph">Nombre del placeholder.</param>
            public void RemovePlaceHolder(string ph)
            {
                placeHolder.Remove(ph);
            }

            /// <summary>
            /// Realiza la combinación de los documentos RTF.
            /// <param name="removeLastPar">Indica si se debe eliminar el último nodo \par de los documentos insertados en la plantilla.</param>
            /// <returns>Devuelve el árbol RTF resultado de la fusión.</returns>
            /// </summary>
            public RtfTree Merge(bool removeLastPar)
            {
                //Indicativo de eliminación del último nodo \par para documentos insertados
                this.removeLastPar = removeLastPar;

                //Se obtiene el grupo principal del árbol
                RtfTreeNode parentNode = baseRtfDoc.MainGroup;

                //Si el documento tiene grupo principal
                if (parentNode != null)
                {
                    //Se analiza el texto del documento en busca de parámetros de reemplazo y se combinan los documentos
                    analizeTextContent(parentNode);
                }

                return baseRtfDoc;
            }

            /// <summary>
            /// Realiza la combinación de los documentos RTF.
            /// <returns>Devuelve el árbol RTF resultado de la fusión.</returns>
            /// </summary>
            public RtfTree Merge()
            {
                return Merge(true);
            }

            #endregion

            #region Propiedades

            /// <summary>
            /// Devuelve la lista de parámetros de sustitución con el formato: [string, RtfTree]
            /// </summary>
            public Dictionary<string, RtfTree> Placeholders
            {
                get
                {
                    return placeHolder;
                }
            }

            /// <summary>
            /// Obtiene o establece el árbol RTF del documento plantilla.
            /// </summary>
            public RtfTree Template
            {
                get
                {
                    return baseRtfDoc;
                }
                set
                {
                    baseRtfDoc = value;
                }
            }

            #endregion

            #region Métodos Privados

            /// <summary>
            /// Analiza el texto del documento en busca de parámetros de reemplazo y combina los documentos.
            /// </summary>
            /// <param name="parentNode">Nodo del árbol a procesar.</param>
            private void analizeTextContent(RtfTreeNode parentNode)
            {
                RtfTree docToInsert = null;
                int indPH;

                //Si el nodo es de tipo grupo y contiene nodos hijos
                if (parentNode != null && parentNode.HasChildNodes())
                {
                    //Se recorren todos los nodos hijos
                    for (int iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
                    {
                        //Nodo actual
                        RtfTreeNode currNode = parentNode.ChildNodes[iNdIndex];

                        //Si el nodo actual es de tipo Texto se buscan etiquetas a reemplazar
                        if (currNode.NodeType == RtfNodeType.Text)
                        {
                            docToInsert = null;

                            //Se recorren todas las etiquetas configuradas
                            foreach (string ph in placeHolder.Keys)
                            {
                                //Se busca la siguiente ocurrencia de la etiqueta actual
                                indPH = currNode.NodeKey.IndexOf(ph);

                                //Si se ha encontrado una etiqueta
                                if (indPH != -1)
                                {
                                    //Se recupera el árbol a insertar en la etiqueta actual
                                    docToInsert = placeHolder[ph].CloneTree();

                                    //Se inserta el nuevo árbol en el árbol base
                                    mergeCore(parentNode, iNdIndex, docToInsert, ph, indPH);

                                    //Como puede que el nodo actual haya cambiado decrementamos el índice
                                    //y salimos del bucle para analizarlo de nuevo
                                    iNdIndex--;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            //Si el nodo actual tiene hijos se analizan los nodos hijos
                            if (currNode.HasChildNodes())
                            {
                                analizeTextContent(currNode);
                            }
                        }
                    }
                }
            }

            /// <summary>
            /// Inserta un nuevo árbol en el lugar de una etiqueta de texto del árbol base.
            /// </summary>
            /// <param name="parentNode">Nodo de tipo grupo que se está procesando.</param>
            /// <param name="iNdIndex">Índice (dentro del grupo padre) del nodo texto que se está procesando.</param>
            /// <param name="docToInsert">Nuevo árbol RTF a insertar.</param>
            /// <param name="strCompletePlaceholder">Texto del la etiqueta que se va a reemplazar.</param>
            /// <param name="intPlaceHolderNodePos">Posición de la etiqueta que se va a reemplazar dentro del nodo texto que se está procesando.</param>
            private void mergeCore(RtfTreeNode parentNode, int iNdIndex, RtfTree docToInsert, string strCompletePlaceholder, int intPlaceHolderNodePos)
            {
                //Si el documento a insertar no está vacío
                if (docToInsert.RootNode.HasChildNodes())
                {
                    int currentIndex = iNdIndex + 1;

                    //Se combinan las tablas de colores y se ajustan los colores del documento a insertar
                    mainAdjustColor(docToInsert);

                    //Se combinan las tablas de fuentes y se ajustan las fuentes del documento a insertar
                    mainAdjustFont(docToInsert);

                    //Se elimina la información de cabecera del documento a insertar (colores, fuentes, info, ...)
                    cleanToInsertDoc(docToInsert);

                    //Si el documento a insertar tiene contenido
                    if (docToInsert.RootNode.FirstChild.HasChildNodes())
                    {
                        //Se inserta el documento nuevo en el árbol base
                        execMergeDoc(parentNode, docToInsert, currentIndex);
                    }

                    //Si la etiqueta no está al final del nodo texto:
                    //Se inserta un nodo de texto con el resto del texto original (eliminando la etiqueta)
                    if (parentNode.ChildNodes[iNdIndex].NodeKey.Length != (intPlaceHolderNodePos + strCompletePlaceholder.Length))
                    {
                        //Se inserta un nodo de texto con el resto del texto original (eliminando la etiqueta)
                        string remText = 
                            parentNode.ChildNodes[iNdIndex].NodeKey.Substring(
                                parentNode.ChildNodes[iNdIndex].NodeKey.IndexOf(strCompletePlaceholder) + strCompletePlaceholder.Length);

                        parentNode.InsertChild(currentIndex + 1, new RtfTreeNode(RtfNodeType.Text, remText, false, 0));
                    }

                    //Si la etiqueta reemplazada estaba al principio del nodo de texto eliminamos el nodo
                    //original porque ya no es necesario
                    if (intPlaceHolderNodePos == 0)
                    {
                        parentNode.RemoveChild(iNdIndex);
                    }
                    //En otro caso lo sustituimos por el texto previo a la etiqueta
                    else  
                    {
                        parentNode.ChildNodes[iNdIndex].NodeKey = 
                            parentNode.ChildNodes[iNdIndex].NodeKey.Substring(0, intPlaceHolderNodePos);
                    }
                }
            }

            /// <summary>
            /// Obtiene el código de la fuente pasada como parámetro, insertándola en la tabla de fuentes si es necesario.
            /// </summary>
            /// <param name="fontDestTbl">Tabla de fuentes resultante.</param>
            /// <param name="sFontName">Fuente buscada.</param>
            /// <returns></returns>
            private int getFontID(ref RtfFontTable fontDestTbl, string sFontName)
            {
                int iExistingFontID = -1;

                if ((iExistingFontID = fontDestTbl.IndexOf(sFontName)) == -1)
                {
                    fontDestTbl.AddFont(sFontName);
                    iExistingFontID = fontDestTbl.IndexOf(sFontName);

                    RtfNodeCollection nodeListToInsert = baseRtfDoc.RootNode.SelectNodes("fonttbl");

                    RtfTreeNode ftFont = new RtfTreeNode(RtfNodeType.Group);
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "f", true, iExistingFontID));
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "fnil", false, 0));
                    ftFont.AppendChild(new RtfTreeNode(RtfNodeType.Text, sFontName + ";", false, 0));
                    
                    nodeListToInsert[0].ParentNode.AppendChild(ftFont);
                }

                return iExistingFontID;
            }

            /// <summary>
            /// Obtiene el código del color pasado como parámetro, insertándolo en la tabla de colores si es necesario.
            /// </summary>
            /// <param name="colorDestTbl">Tabla de colores resultante.</param>
            /// <param name="iColorName">Color buscado.</param>
            /// <returns></returns>
            private int getColorID(RtfColorTable colorDestTbl, Color iColorName)
            {
                int iExistingColorID;

                if ((iExistingColorID = colorDestTbl.IndexOf(iColorName)) == -1)
                {
                    iExistingColorID = colorDestTbl.Count;
                    colorDestTbl.AddColor(iColorName);

                    RtfNodeCollection nodeListToInsert = baseRtfDoc.RootNode.SelectNodes("colortbl");

                    nodeListToInsert[0].ParentNode.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "red", true, iColorName.R));
                    nodeListToInsert[0].ParentNode.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "green", true, iColorName.G));
                    nodeListToInsert[0].ParentNode.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "blue", true, iColorName.B));
                    nodeListToInsert[0].ParentNode.AppendChild(new RtfTreeNode(RtfNodeType.Text, ";", false, 0));
                }

                return iExistingColorID;
            }

            /// <summary>
            /// Ajusta las fuentes del documento a insertar.
            /// </summary>
            /// <param name="docToInsert">Documento a insertar.</param>
            private void mainAdjustFont(RtfTree docToInsert)
            {
                RtfFontTable fontDestTbl = baseRtfDoc.GetFontTable();
                RtfFontTable fontToCopyTbl = docToInsert.GetFontTable();

                adjustFontRecursive(docToInsert.RootNode, fontDestTbl, fontToCopyTbl);
            }

            /// <summary>
            /// Ajusta las fuentes del documento a insertar.
            /// </summary>
            /// <param name="parentNode">Nodo grupo que se está procesando.</param>
            /// <param name="fontDestTbl">Tabla de fuentes resultante.</param>
            /// <param name="fontToCopyTbl">Tabla de fuentes del documento a insertar.</param>
            private void adjustFontRecursive(RtfTreeNode parentNode, RtfFontTable fontDestTbl, RtfFontTable fontToCopyTbl)
            {
                if (parentNode != null && parentNode.HasChildNodes())
                {
                    for (int iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
                    {
                        if (parentNode.ChildNodes[iNdIndex].NodeType == RtfNodeType.Keyword &&
                            (parentNode.ChildNodes[iNdIndex].NodeKey == "f" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "stshfdbch" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "stshfloch" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "stshfhich" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "stshfbi" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "deff" ||
                            parentNode.ChildNodes[iNdIndex].NodeKey == "af") &&
                            parentNode.ChildNodes[iNdIndex].HasParameter)
                        {
                            parentNode.ChildNodes[iNdIndex].Parameter = getFontID(ref fontDestTbl, fontToCopyTbl[parentNode.ChildNodes[iNdIndex].Parameter]);
                        }

                        adjustFontRecursive(parentNode.ChildNodes[iNdIndex], fontDestTbl, fontToCopyTbl);
                    }
                }
            }

            /// <summary>
            /// Ajusta los colores del documento a insertar.
            /// </summary>
            /// <param name="docToInsert">Documento a insertar.</param>
            private void mainAdjustColor(RtfTree docToInsert)
            {
                RtfColorTable colorDestTbl = baseRtfDoc.GetColorTable();
                RtfColorTable colorToCopyTbl = docToInsert.GetColorTable();

                adjustColorRecursive(docToInsert.RootNode, colorDestTbl, colorToCopyTbl);
            }

            /// <summary>
            /// Ajusta los colores del documento a insertar.
            /// </summary>
            /// <param name="parentNode">Nodo grupo que se está procesando.</param>
            /// <param name="colorDestTbl">Tabla de colores resultante.</param>
            /// <param name="colorToCopyTbl">Tabla de colores del documento a insertar.</param>
            private void adjustColorRecursive(RtfTreeNode parentNode, RtfColorTable colorDestTbl, RtfColorTable colorToCopyTbl)
            {
                if (parentNode != null && parentNode.HasChildNodes())
                {
                    for (int iNdIndex = 0; iNdIndex < parentNode.ChildNodes.Count; iNdIndex++)
                    {
                        if (parentNode.ChildNodes[iNdIndex].NodeType == RtfNodeType.Keyword &&
                            (parentNode.ChildNodes[iNdIndex].NodeKey == "cf" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "cb" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "pncf" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "brdrcf" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "cfpat" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "cbpat" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "clcfpatraw" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "clcbpatraw" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "ulc" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "chcfpat" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "chcbpat" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "highlight" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "clcbpat" ||
                             parentNode.ChildNodes[iNdIndex].NodeKey == "clcfpat") &&
                             parentNode.ChildNodes[iNdIndex].HasParameter)
                        {
                            parentNode.ChildNodes[iNdIndex].Parameter = getColorID(colorDestTbl, colorToCopyTbl[parentNode.ChildNodes[iNdIndex].Parameter]);
                        }

                        adjustColorRecursive(parentNode.ChildNodes[iNdIndex], colorDestTbl, colorToCopyTbl);
                    }
                }
            }

            /// <summary>
            /// Inserta el nuevo árbol en el árbol base (como un nuevo grupo) eliminando toda la cabecera del documento insertado.
            /// </summary>
            /// <param name="parentNode">Grupo base en el que se insertará el nuevo arbol.</param>
            /// <param name="treeToCopyParent">Nuevo árbol a insertar.</param>
            /// <param name="intCurrIndex">Índice en el que se insertará el nuevo árbol dentro del grupo base.</param>
            private void execMergeDoc(RtfTreeNode parentNode, RtfTree treeToCopyParent, int intCurrIndex)
            {
                //Se busca el primer "\pard" del documento (comienzo del texto)
                RtfTreeNode nodePard = treeToCopyParent.RootNode.FirstChild.SelectSingleChildNode("pard");

                //Se obtiene el índice del nodo dentro del principal
                int indPard = treeToCopyParent.RootNode.FirstChild.ChildNodes.IndexOf(nodePard);

                //Se crea el nuevo grupo
                RtfTreeNode newGroup = new RtfTreeNode(RtfNodeType.Group);

                //Se resetean las opciones de párrafo y fuente
                newGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "pard", false, 0));
                newGroup.AppendChild(new RtfTreeNode(RtfNodeType.Keyword, "plain", false, 0));

                //Se inserta cada nodo hijo del documento nuevo en el documento base
                for (int i = indPard + 1; i < treeToCopyParent.RootNode.FirstChild.ChildNodes.Count; i++)
                {
                    RtfTreeNode newNode = 
                        treeToCopyParent.RootNode.FirstChild.ChildNodes[i].CloneNode();

                    newGroup.AppendChild(newNode);
                }

                //Se inserta el nuevo grupo con el nuevo documento
                parentNode.InsertChild(intCurrIndex, newGroup);
            }

            /// <summary>
            /// Elimina los elementos no deseados del documento a insertar, por ejemplo los nodos "\par" finales.
            /// </summary>
            /// <param name="docToInsert">Documento a insertar.</param>
            private void cleanToInsertDoc(RtfTree docToInsert)
            {
                //Borra el último "\par" del documento si existe
                RtfTreeNode lastNode = docToInsert.RootNode.FirstChild.LastChild;

                if (removeLastPar)
                {
                    if (lastNode.NodeType == RtfNodeType.Keyword && lastNode.NodeKey == "par")
                    {
                        docToInsert.RootNode.FirstChild.RemoveChild(lastNode);
                    }
                }
            }

            #endregion
        }
    }
}
