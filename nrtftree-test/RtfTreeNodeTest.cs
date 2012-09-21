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
        public void NodeText()
        {
            RtfTree tree = new RtfTree();

            int res = tree.LoadRtfFile("..\\..\\testdocs\\testdoc4.rtf");

            RtfTreeNode node1 = tree.MainGroup.SelectSingleGroup("ul");
            RtfTreeNode node2 = tree.MainGroup.SelectSingleGroup("cf");

            Assert.That(node1.Text, Is.EqualTo("underline1"));
            Assert.That(node2.Text, Is.EqualTo("blue1 luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et italic1 dis parturient montes, nascetur ridiculus mus."));
        }
    }
}
