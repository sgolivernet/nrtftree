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
 * Class:		LoadRtfTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using System.IO;
using NUnit.Framework;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class LoadRtfTest
    {
        [OneTimeSetUp]
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
        public void LoadSimpleDocFromFile()
        {
            RtfTree tree = new RtfTree();

            int res = tree.LoadRtfFile("testdocs\\testdoc1.rtf");

            //StreamWriter sw = null;
            //sw = new StreamWriter("testdocs\\result1-1.txt");
            //sw.Write(tree.ToString());
            //sw.Flush();
            //sw.Close();
            //sw = new StreamWriter("testdocs\\result1-2.txt");
            //sw.Write(tree.ToStringEx());
            //sw.Flush();
            //sw.Close();
            //sw = new StreamWriter("testdocs\\rtf1.txt");
            //sw.Write(tree.Rtf);
            //sw.Flush();
            //sw.Close();
            //sw = new StreamWriter("testdocs\\text1.txt");
            //sw.Write(tree.Text);
            //sw.Flush();
            //sw.Close();

            StreamReader sr = null;

            sr = new StreamReader("testdocs\\result1-1.txt");
            string strTree1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\result1-2.txt");
            string strTree2 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf1.txt");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\text1.txt");
            string text1 = sr.ReadToEnd();
            sr.Close();

            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.False);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        }

        [Test]
        public void LoadImageDocFromFile()
        {
            RtfTree tree = new RtfTree();

            int res = tree.LoadRtfFile("testdocs\\testdoc3.rtf");

            //StreamWriter sw = null;
            //sw = new StreamWriter("testdocs\\rtf5.txt");
            //sw.Write(tree.Rtf);
            //sw.Flush();
            //sw.Close();

            StreamReader sr = null;

            sr = new StreamReader("testdocs\\rtf5.txt");
            string rtf5 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\text2.txt");
            string text2 = sr.ReadToEnd();
            sr.Close();

            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.False);
            Assert.That(tree.Rtf, Is.EqualTo(rtf5));
            Assert.That(tree.Text, Is.EqualTo(text2));
        }

        [Test]
        public void LoadSimpleDocMergeSpecialFromFile()
        {
            RtfTree tree = new RtfTree();

            tree.MergeSpecialCharacters = true;

            int res = tree.LoadRtfFile("testdocs\\testdoc1.rtf");

            //StreamWriter sw = null;
            //sw = new StreamWriter("testdocs\\result1-3.txt");
            //sw.Write(tree.ToString());
            //sw.Flush();
            //sw.Close();
            //sw = new StreamWriter("testdocs\\result1-4.txt");
            //sw.Write(tree.ToStringEx());
            //sw.Flush();
            //sw.Close();

            StreamReader sr = null;

            sr = new StreamReader("testdocs\\result1-3.txt");
            string strTree1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\result1-4.txt");
            string strTree2 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf1.txt");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\text1.txt");
            string text1 = sr.ReadToEnd();
            sr.Close();

            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.True);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        }

        [Test]
        public void LoadSimpleDocFromString()
        {
            RtfTree tree = new RtfTree();

            StreamReader sr = new StreamReader("testdocs\\testdoc1.rtf");
            string strDoc = sr.ReadToEnd();
            sr.Close();

            int res = tree.LoadRtfText(strDoc);

            sr = new StreamReader("testdocs\\result1-1.txt");
            string strTree1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\result1-2.txt");
            string strTree2 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf1.txt");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\text1.txt");
            string text1 = sr.ReadToEnd();
            sr.Close();

            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.False);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        }

        [Test]
        public void LoadSimpleDocMergeSpecialFromString()
        {
            RtfTree tree = new RtfTree();
            tree.MergeSpecialCharacters = true;

            StreamReader sr = new StreamReader("testdocs\\testdoc1.rtf");
            string strDoc = sr.ReadToEnd();
            sr.Close();

            int res = tree.LoadRtfText(strDoc);

            sr = new StreamReader("testdocs\\result1-3.txt");
            string strTree1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\result1-4.txt");
            string strTree2 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf1.txt");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\text1.txt");
            string text1 = sr.ReadToEnd();
            sr.Close();

            Assert.That(res, Is.EqualTo(0));
            Assert.That(tree.MergeSpecialCharacters, Is.True);
            Assert.That(tree.ToString(), Is.EqualTo(strTree1));
            Assert.That(tree.ToStringEx(), Is.EqualTo(strTree2));
            Assert.That(tree.Rtf, Is.EqualTo(rtf1));
            Assert.That(tree.Text, Is.EqualTo(text1));
        }
    }
}
