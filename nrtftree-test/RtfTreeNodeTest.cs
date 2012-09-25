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
 * Class:		RtfTreeNodeTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using NUnit.Framework;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class RtfTreeNodeTest
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
        public void AddChildToEmptyNode()
        {
            RtfTreeNode node = new RtfTreeNode();

            Assert.That(node.ChildNodes, Is.Null);

            RtfTreeNode childNode = new RtfTreeNode();
            node.InsertChild(0, childNode);

            Assert.That(node.ChildNodes, Is.Not.Null);
            Assert.That(node.ChildNodes[0], Is.SameAs(childNode));
            Assert.That(childNode.ChildNodes, Is.Null);

            RtfTreeNode anotherChildNode = new RtfTreeNode();
            childNode.AppendChild(anotherChildNode);

            Assert.That(childNode.ChildNodes, Is.Not.Null);
            Assert.That(childNode.ChildNodes[0], Is.SameAs(anotherChildNode));
        }

        [Test]
        public void StringRepresentation()
        {
            RtfTreeNode node = new RtfTreeNode(RtfNodeType.Keyword, "b", true, 3);
            RtfTreeNode node2 = new RtfTreeNode(RtfNodeType.Root);

            Assert.That(node.ToString(), Is.EqualTo("[Keyword, b, True, 3]"));
            Assert.That(node2.ToString(), Is.EqualTo("[Root, , False, 0]"));
        }

        [Test]
        public void TextExtraction()
        {
            RtfTree tree = new RtfTree();

            int res = tree.LoadRtfFile("..\\..\\testdocs\\testdoc4.rtf");

            RtfTreeNode simpleGroup = tree.MainGroup.SelectSingleGroup("ul");
            RtfTreeNode nestedGroups = tree.MainGroup.SelectSingleGroup("cf");
            RtfTreeNode keyword = tree.MainGroup.SelectSingleChildNode("b");
            RtfTreeNode control = tree.MainGroup.SelectSingleChildNode("'");
            RtfTreeNode root = tree.RootNode;

            Assert.That(simpleGroup.Text, Is.EqualTo("underline1"));
            Assert.That(nestedGroups.Text, Is.EqualTo("blue1 luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et italic1 dis parturient montes, nascetur ridiculus mus."));
            Assert.That(keyword.Text, Is.EqualTo(""));
            Assert.That(control.Text, Is.EqualTo("é"));
            Assert.That(root.Text, Is.EqualTo(""));

            Assert.That(simpleGroup.RawText, Is.EqualTo("underline1"));
            Assert.That(nestedGroups.RawText, Is.EqualTo("blue1 luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et italic1 dis parturient montes, nascetur ridiculus mus."));
            Assert.That(keyword.RawText, Is.EqualTo(""));
            Assert.That(control.RawText, Is.EqualTo("é"));
            Assert.That(root.RawText, Is.EqualTo(""));

            RtfTreeNode fontsGroup = tree.MainGroup.SelectSingleGroup("fonttbl");
            RtfTreeNode generatorGroup = tree.MainGroup.SelectSingleGroup("*");

            Assert.That(fontsGroup.Text, Is.EqualTo(""));
            Assert.That(generatorGroup.Text, Is.EqualTo(""));

            Assert.That(fontsGroup.RawText, Is.EqualTo("Times New Roman;Arial;Arial;"));
            Assert.That(generatorGroup.RawText, Is.EqualTo("Msftedit 5.41.15.1515;"));
        }

        [Test]
        public void TextExtractionSpecial()
        {
            RtfTree tree = new RtfTree();

            int res = tree.LoadRtfFile("..\\..\\testdocs\\testdoc5.rtf");

            Assert.That(tree.Text, Is.EqualTo("Esto es una ‘prueba’\r\n\t y otra “prueba” y otra—prueba." + Environment.NewLine));
            Assert.That(tree.MainGroup.Text, Is.EqualTo("Esto es una ‘prueba’\r\n\t y otra “prueba” y otra—prueba." + Environment.NewLine));
            Assert.That(tree.MainGroup.RawText, Is.EqualTo("Arial;Msftedit 5.41.15.1515;Esto es una ‘prueba’\r\n\t y otra “prueba” y otra—prueba." + Environment.NewLine));
        }
    }
}
