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
 * Class:		NavigationTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using NUnit.Framework;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class NavigationTest
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
        public void EmptyNodeNavigation()
        {
            RtfTreeNode node = new RtfTreeNode();

            Assert.That(node.Tree, Is.Null);
            Assert.That(node.RootNode, Is.Null);
            Assert.That(node.ParentNode, Is.Null);

            Assert.That(node.NextSibling, Is.Null);
            Assert.That(node.PreviousSibling, Is.Null);

            Assert.That(node.ChildNodes, Is.Null);

            Assert.That(node.FirstChild, Is.Null);
            Assert.That(node.LastChild, Is.Null);
        }

        [Test]
        public void TypeNodeNavigation()
        {
            RtfTreeNode node = new RtfTreeNode(RtfNodeType.Keyword);

            Assert.That(node.Tree, Is.Null);
            Assert.That(node.RootNode, Is.Null);
            Assert.That(node.ParentNode, Is.Null);

            Assert.That(node.NextSibling, Is.Null);
            Assert.That(node.PreviousSibling, Is.Null);

            Assert.That(node.ChildNodes, Is.Null);

            Assert.That(node.FirstChild, Is.Null);
            Assert.That(node.LastChild, Is.Null);
        }

        [Test]
        public void InitNodeNavigation()
        {
            RtfTreeNode node = new RtfTreeNode(RtfNodeType.Keyword, "rtf", true, 99);

            Assert.That(node.Tree, Is.Null);
            Assert.That(node.RootNode, Is.Null);
            Assert.That(node.ParentNode, Is.Null);

            Assert.That(node.NextSibling, Is.Null);
            Assert.That(node.PreviousSibling, Is.Null);

            Assert.That(node.ChildNodes, Is.Null);

            Assert.That(node.FirstChild, Is.Null);
            Assert.That(node.LastChild, Is.Null);
        }

        [Test]
        public void EmptyTreeNavigation()
        {
            RtfTree tree = new RtfTree();

            Assert.That(tree.RootNode, Is.Not.Null);
            Assert.That(tree.RootNode.Tree, Is.SameAs(tree));
            Assert.That(tree.MainGroup, Is.Null);
        }

        [Test]
        public void SimpleTreeNavigation()
        {
            //Creación de un árbol sencillo

            RtfTree tree = new RtfTree();

            RtfTreeNode mainGroup = new RtfTreeNode(RtfNodeType.Group);
            RtfTreeNode rtfNode = new RtfTreeNode(RtfNodeType.Keyword, "rtf", true, 0);
            mainGroup.AppendChild(rtfNode);

            RtfTreeNode newGroup = new RtfTreeNode(RtfNodeType.Group);
            RtfTreeNode node1 = new RtfTreeNode(RtfNodeType.Keyword, "ul", false, 0);
            RtfTreeNode node2 = new RtfTreeNode(RtfNodeType.Text, "Test", false, 0);
            RtfTreeNode node3 = new RtfTreeNode(RtfNodeType.Text, "ulnone", false, 0);

            newGroup.AppendChild(node1);
            newGroup.AppendChild(node2);
            newGroup.AppendChild(node3);

            mainGroup.AppendChild(newGroup);

            tree.RootNode.AppendChild(mainGroup);

            //Navegación básica: tree
            Assert.That(tree.RootNode, Is.Not.Null);
            Assert.That(tree.MainGroup, Is.SameAs(mainGroup));

            //Navegación básica: newGroup
            Assert.That(newGroup.Tree, Is.SameAs(tree));
            Assert.That(newGroup.ParentNode, Is.SameAs(mainGroup));
            Assert.That(newGroup.RootNode, Is.SameAs(tree.RootNode));
            Assert.That(newGroup.ChildNodes, Is.Not.Null);
            Assert.That(newGroup[1], Is.SameAs(node2));
            Assert.That(newGroup.ChildNodes[1], Is.SameAs(node2));
            Assert.That(newGroup["ul"], Is.SameAs(node1));
            Assert.That(newGroup.FirstChild, Is.SameAs(node1));
            Assert.That(newGroup.LastChild, Is.SameAs(node3));
            Assert.That(newGroup.PreviousSibling, Is.SameAs(rtfNode));
            Assert.That(newGroup.NextSibling, Is.Null);
            Assert.That(newGroup.Index, Is.EqualTo(1));

            //Navegación básica: nodo2
            Assert.That(node2.Tree, Is.SameAs(tree));
            Assert.That(node2.ParentNode, Is.SameAs(newGroup));
            Assert.That(node2.RootNode, Is.SameAs(tree.RootNode));
            Assert.That(node2.ChildNodes, Is.Null);
            Assert.That(node2[1], Is.Null);
            Assert.That(node2["ul"], Is.Null);
            Assert.That(node2.FirstChild, Is.Null);
            Assert.That(node2.LastChild, Is.Null);
            Assert.That(node2.PreviousSibling, Is.SameAs(node1));
            Assert.That(node2.NextSibling, Is.SameAs(node3));
            Assert.That(node2.Index, Is.EqualTo(1));
        }

        [Test]
        public void AdjacentNodes()
        {
            //Creación de un árbol sencillo

            RtfTree tree = new RtfTree();

            RtfTreeNode mainGroup = new RtfTreeNode(RtfNodeType.Group);
            RtfTreeNode rtfNode = new RtfTreeNode(RtfNodeType.Keyword, "rtf", true, 0);
            mainGroup.AppendChild(rtfNode);

            RtfTreeNode newGroup = new RtfTreeNode(RtfNodeType.Group);
            RtfTreeNode node1 = new RtfTreeNode(RtfNodeType.Keyword, "ul", false, 0);
            RtfTreeNode node2 = new RtfTreeNode(RtfNodeType.Text, "Test", false, 0);
            RtfTreeNode node3 = new RtfTreeNode(RtfNodeType.Keyword, "ulnone", false, 0);

            newGroup.AppendChild(node1);
            newGroup.AppendChild(node2);
            newGroup.AppendChild(node3);

            mainGroup.AppendChild(newGroup);

            tree.RootNode.AppendChild(mainGroup);

            RtfTreeNode node4 = new RtfTreeNode(RtfNodeType.Text, "fin", false, 0);

            mainGroup.AppendChild(node4);

            Assert.That(tree.RootNode.NextNode, Is.SameAs(mainGroup));
            Assert.That(mainGroup.NextNode, Is.SameAs(rtfNode));
            Assert.That(rtfNode.NextNode, Is.SameAs(newGroup));
            Assert.That(newGroup.NextNode, Is.SameAs(node1));
            Assert.That(node1.NextNode, Is.SameAs(node2));
            Assert.That(node2.NextNode, Is.SameAs(node3));
            Assert.That(node3.NextNode, Is.SameAs(node4));
            Assert.That(node4.NextNode, Is.Null);

            Assert.That(node4.PreviousNode, Is.SameAs(node3));
            Assert.That(node3.PreviousNode, Is.SameAs(node2));
            Assert.That(node2.PreviousNode, Is.SameAs(node1));
            Assert.That(node1.PreviousNode, Is.SameAs(newGroup));
            Assert.That(newGroup.PreviousNode, Is.SameAs(rtfNode));
            Assert.That(rtfNode.PreviousNode, Is.SameAs(mainGroup));
            Assert.That(mainGroup.PreviousNode, Is.SameAs(tree.RootNode));
            Assert.That(tree.RootNode.PreviousNode, Is.Null);
        }
    }
}
