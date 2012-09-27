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
 * Class:		MergeTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using System.IO;
using NUnit.Framework;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class MergeTest
    {
        [TestFixtureSetUp]
        public void InitTestFixture()
        {
            ;
        }

        [SetUp]
        public void InitTest()
        {
            ;
        }

        [Test]
        public void MergeDocumentsFile()
        {
            RtfMerger merger = new RtfMerger("..\\..\\testdocs\\merge-template.rtf");
            merger.AddPlaceHolder("$doc1$", "..\\..\\testdocs\\merge-doc1.rtf");
            merger.AddPlaceHolder("$doc2$", "..\\..\\testdocs\\merge-doc2.rtf");

            Assert.That(merger.Placeholders.Count, Is.EqualTo(2));

            merger.AddPlaceHolder("$doc3$", "..\\..\\testdocs\\merge-doc2.rtf");

            Assert.That(merger.Placeholders.Count, Is.EqualTo(3));

            merger.RemovePlaceHolder("$doc3$");

            Assert.That(merger.Placeholders.Count, Is.EqualTo(2));

            RtfTree tree = merger.Merge();
            tree.SaveRtf("..\\..\\testdocs\\merge-result-1.rtf");

            StreamReader sr = null;
            sr = new StreamReader("..\\..\\testdocs\\merge-result-1.rtf");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("..\\..\\testdocs\\rtf3.txt");
            string rtf3 = sr.ReadToEnd();
            sr.Close();

            Assert.That(rtf1, Is.EqualTo(rtf3));
        }

        [Test]
        public void MergeDocumentsInMemory()
        {
            RtfMerger merger = new RtfMerger();

            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\merge-template.rtf");

            merger.Template = tree;

            RtfTree ph1 = new RtfTree();
            ph1.LoadRtfFile("..\\..\\testdocs\\merge-doc1.rtf");

            RtfTree ph2 = new RtfTree();
            ph2.LoadRtfFile("..\\..\\testdocs\\merge-doc2.rtf");

            merger.AddPlaceHolder("$doc1$", ph1);
            merger.AddPlaceHolder("$doc2$", ph2);

            Assert.That(merger.Placeholders.Count, Is.EqualTo(2));

            RtfTree ph3 = new RtfTree();
            ph3.LoadRtfFile("..\\..\\testdocs\\merge-doc2.rtf");

            merger.AddPlaceHolder("$doc3$", ph3);

            Assert.That(merger.Placeholders.Count, Is.EqualTo(3));

            merger.RemovePlaceHolder("$doc3$");

            Assert.That(merger.Placeholders.Count, Is.EqualTo(2));

            RtfTree resTree = merger.Merge();
            resTree.SaveRtf("..\\..\\testdocs\\merge-result-2.rtf");

            StreamReader sr = null;
            sr = new StreamReader("..\\..\\testdocs\\merge-result-2.rtf");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("..\\..\\testdocs\\rtf3.txt");
            string rtf3 = sr.ReadToEnd();
            sr.Close();

            Assert.That(rtf1, Is.EqualTo(rtf3));
        }
    }
}
