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
 * Version:     v0.3.0
 * Date:		01/05/2009
 * Copyright:   2006-2009 Salvador Gomez
 * E-mail:      sgoliver.net@gmail.com
 * Home Page:	http://www.sgoliver.net
 * SF Project:	http://nrtftree.sourceforge.net
 *				http://sourceforge.net/projects/nrtftree
 * Class:		HeaderSectionTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;
using System.IO;
using NUnit.Framework;
using System.Drawing;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class HeaderSectionsTest
    {
        RtfTree tree = null;

        [TestFixtureSetUp]
        public void InitTestFixture()
        {
            tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc2.rtf");
        }

        [SetUp]
        public void InitTest()
        {
            ;
        }

        [Test]
        public void FontTableTest()
        {
            RtfFontTable fontTable = tree.GetFontTable();

            Assert.That(fontTable.Count, Is.EqualTo(3));
            Assert.That(fontTable[0], Is.EqualTo("Times New Roman"));
            Assert.That(fontTable[1], Is.EqualTo("Arial"));
            Assert.That(fontTable[2], Is.EqualTo("Arial"));

            Assert.That(fontTable.IndexOf("Times New Roman"), Is.EqualTo(0));
            Assert.That(fontTable.IndexOf("Arial"), Is.EqualTo(1));
            Assert.That(fontTable.IndexOf("nofont"), Is.EqualTo(-1));
        }

        [Test]
        public void ColorTableTest()
        {
            RtfColorTable colorTable = tree.GetColorTable();

            Assert.That(colorTable.Count, Is.EqualTo(3));
            Assert.That(colorTable[0], Is.EqualTo(Color.FromArgb(0,0,0)));
            Assert.That(colorTable[1], Is.EqualTo(Color.FromArgb(0, 0, 128)));
            Assert.That(colorTable[2], Is.EqualTo(Color.FromArgb(255, 0, 0)));

            Assert.That(colorTable.IndexOf(Color.FromArgb(0, 0, 0)), Is.EqualTo(0));
            Assert.That(colorTable.IndexOf(Color.FromArgb(0, 0, 128)), Is.EqualTo(1));
            Assert.That(colorTable.IndexOf(Color.FromArgb(255, 0, 0)), Is.EqualTo(2));
        }

        [Test]
        public void StyleSheetTableTest()
        {
            RtfStyleSheetTable styleTable = tree.GetStyleSheetTable();

            Assert.That(styleTable.Count, Is.EqualTo(7));

            Assert.That(styleTable[0].Index, Is.EqualTo(0));
            Assert.That(styleTable[0].Type, Is.EqualTo(RtfStyleSheetType.Paragraph));
            Assert.That(styleTable[0].Name, Is.EqualTo("Normal"));
            Assert.That(styleTable[0].Next, Is.EqualTo(0));
            Assert.That(styleTable[0].Formatting.Count, Is.EqualTo(25));

            Assert.That(styleTable[1].Index, Is.EqualTo(1));
            Assert.That(styleTable[1].Type, Is.EqualTo(RtfStyleSheetType.Paragraph));
            Assert.That(styleTable[1].Name, Is.EqualTo("heading 1"));
            Assert.That(styleTable[1].Next, Is.EqualTo(0));
            Assert.That(styleTable[1].BasedOn, Is.EqualTo(0));
            Assert.That(styleTable[1].Styrsid, Is.EqualTo(2310575));
            Assert.That(styleTable[1].Formatting.Count, Is.EqualTo(33));

            Assert.That(styleTable[10].Index, Is.EqualTo(10));
            Assert.That(styleTable[10].Type, Is.EqualTo(RtfStyleSheetType.Character));
            Assert.That(styleTable[10].Name, Is.EqualTo("Default Paragraph Font"));
            Assert.That(styleTable[10].Additive, Is.EqualTo(true));
            Assert.That(styleTable[10].SemiHidden, Is.EqualTo(true));
            Assert.That(styleTable[10].Formatting.Count, Is.EqualTo(0));

            Assert.That(styleTable[11].Index, Is.EqualTo(11));
            Assert.That(styleTable[11].Type, Is.EqualTo(RtfStyleSheetType.Table));
            Assert.That(styleTable[11].Name, Is.EqualTo("Normal Table"));
            Assert.That(styleTable[11].Next, Is.EqualTo(11));
            Assert.That(styleTable[11].SemiHidden, Is.EqualTo(true));
            Assert.That(styleTable[11].Formatting.Count, Is.EqualTo(44));
        }

        [Test]
        public void InfoGroupTest()
        {
            InfoGroup infoGroup = tree.GetInfoGroup();

            Assert.That(infoGroup.Title, Is.EqualTo("Test NRtfTree Title"));
            Assert.That(infoGroup.Subject, Is.EqualTo("Test NRtfTree Subject"));
            Assert.That(infoGroup.Author, Is.EqualTo("Sgoliver (Author)"));
            Assert.That(infoGroup.Keywords, Is.EqualTo("test;nrtftree;sgoliver"));
            Assert.That(infoGroup.DocComment, Is.EqualTo("This is a test comment."));
            Assert.That(infoGroup.Operator, Is.EqualTo("None"));
            Assert.That(infoGroup.CreationTime, Is.EqualTo(new DateTime(2008, 5, 28, 18, 52, 00)));
            Assert.That(infoGroup.RevisionTime, Is.EqualTo(new DateTime(2009, 6, 29, 20, 23, 00)));
            Assert.That(infoGroup.Version, Is.EqualTo(6));
            Assert.That(infoGroup.EditingTime, Is.EqualTo(4));
            Assert.That(infoGroup.NumberOfPages, Is.EqualTo(1));
            Assert.That(infoGroup.NumberOfWords, Is.EqualTo(12));
            Assert.That(infoGroup.NumberOfChars, Is.EqualTo(59));
            Assert.That(infoGroup.Manager, Is.EqualTo("Sgoliver (Admin)"));
            Assert.That(infoGroup.Company, Is.EqualTo("www.sgoliver.net"));
            Assert.That(infoGroup.Category, Is.EqualTo("Demos (Category)"));
            Assert.That(infoGroup.InternalVersion, Is.EqualTo(24579));

            Assert.That(infoGroup.Comment, Is.EqualTo(""));
            Assert.That(infoGroup.HlinkBase, Is.EqualTo(""));
            Assert.That(infoGroup.Id, Is.EqualTo(-1));
            Assert.That(infoGroup.LastPrintTime, Is.EqualTo(DateTime.MinValue));
            Assert.That(infoGroup.BackupTime, Is.EqualTo(DateTime.MinValue));

            //StreamWriter sw = new StreamWriter("testdocs\\infogroup.txt");
            //sw.Write(infoGroup.ToString());
            //sw.Flush();
            //sw.Close();

            StreamReader sr = new StreamReader("..\\..\\testdocs\\infogroup.txt");
            string infoString = sr.ReadToEnd();
            sr.Close();

            Assert.That(infoGroup.ToString(), Is.EqualTo(infoString));
        }
    }
}
