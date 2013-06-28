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
 * Class:		ImageNodeTest
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
    public class ImageNodeTest
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
        public void LoadImageNode()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

            ImageNode imgNode = new ImageNode(pictNode);

            Assert.That(imgNode.Height, Is.EqualTo(6615));
            Assert.That(imgNode.Width, Is.EqualTo(7938));

            Assert.That(imgNode.DesiredHeight, Is.EqualTo(3750));
            Assert.That(imgNode.DesiredWidth, Is.EqualTo(4500));

            Assert.That(imgNode.ScaleX, Is.EqualTo(100));
            Assert.That(imgNode.ScaleY, Is.EqualTo(100));

            Assert.That(imgNode.ImageFormat, Is.EqualTo(ImageFormat.Png));
        }

        [Test]
        public void ImageHexData()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

            ImageNode imgNode = new ImageNode(pictNode);

            StreamReader sr = null;

            sr = new StreamReader("..\\..\\testdocs\\imghexdata.txt");
            string hexdata = sr.ReadToEnd();
            sr.Close();

            Assert.That(imgNode.HexData, Is.EqualTo(hexdata));
        }

        [Test]
        public void ImageBinData()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc3.rtf");

            RtfTreeNode pictNode = tree.MainGroup.SelectNodes("pict")[2].ParentNode;

            ImageNode imgNode = new ImageNode(pictNode);

            imgNode.SaveImage("..\\..\\testdocs\\img-result.png", ImageFormat.Jpeg);

            Stream fs1 = new FileStream("..\\..\\testdocs\\img-result.jpg", FileMode.Open);
            Stream fs2 = new FileStream("..\\..\\testdocs\\image1.jpg", FileMode.Open);

            Assert.That(fs1, Is.EqualTo(fs2));
        }
    }
}
