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
 * Class:		ObjectNodeTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;
using System.IO;
using NUnit.Framework;
using System.Drawing;
using System.Drawing.Imaging;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class ObjectNodeTest
    {
        [TestFixtureSetUp]
        public void InitTestFixture()
        {

        }

        [SetUp]
        public void InitTest()
        {
            ;
        }

        [Test]
        public void LoadObjectNode()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode node = tree.MainGroup.SelectSingleNode("object").ParentNode;

            ObjectNode objNode = new ObjectNode(node);

            Assert.That(objNode.ObjectType, Is.EqualTo("objemb"));
            Assert.That(objNode.ObjectClass, Is.EqualTo("Excel.Sheet.8"));
        }

        [Test]
        public void ObjectHexData()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode node = tree.MainGroup.SelectSingleNode("object").ParentNode;

            ObjectNode objNode = new ObjectNode(node);

            StreamReader sr = null;

            sr = new StreamReader("..\\..\\testdocs\\objhexdata.txt");
            string hexdata = sr.ReadToEnd();
            sr.Close();

            Assert.That(objNode.HexData, Is.EqualTo(hexdata));
        }

        [Test]
        public void ObjectBinData()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode node = tree.MainGroup.SelectSingleNode("object").ParentNode;

            ObjectNode objNode = new ObjectNode(node);

            BinaryWriter bw = new BinaryWriter(new FileStream("..\\..\\testdocs\\objbindata-result.dat", FileMode.Create));
            foreach (byte b in objNode.GetByteData())
                bw.Write(b);
            bw.Close();

            FileStream fs1 = new FileStream("..\\..\\testdocs\\objbindata-result.dat", FileMode.Open);
            FileStream fs2 = new FileStream("..\\..\\testdocs\\objbindata.dat", FileMode.Open);

            Assert.That(fs1, Is.EqualTo(fs2));
        }

        [Test]
        public void ResultNode()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode node = tree.MainGroup.SelectSingleNode("object").ParentNode;

            ObjectNode objNode = new ObjectNode(node);

            RtfTreeNode resNode = objNode.ResultNode;

            Assert.That(resNode, Is.SameAs(tree.MainGroup.SelectSingleGroup("object").SelectSingleChildGroup("result")));

            RtfTreeNode pictNode = resNode.SelectSingleNode("pict").ParentNode;
            ImageNode imgNode = new ImageNode(pictNode);

            Assert.That(imgNode.Height, Is.EqualTo(2247));
            Assert.That(imgNode.Width, Is.EqualTo(9320));

            Assert.That(imgNode.DesiredHeight, Is.EqualTo(1274));
            Assert.That(imgNode.DesiredWidth, Is.EqualTo(5284));

            Assert.That(imgNode.ScaleX, Is.EqualTo(100));
            Assert.That(imgNode.ScaleY, Is.EqualTo(100));

            Assert.That(imgNode.ImageFormat, Is.EqualTo(ImageFormat.Emf));
        }
    }
}
