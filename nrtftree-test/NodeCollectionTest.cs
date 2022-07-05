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
 * Class:		NodeCollectionTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using System.IO;
using NUnit.Framework;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class NodeCollectionTest
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
        public void PopulateCollection()
        {
            RtfNodeCollection list1 = new RtfNodeCollection();
            RtfNodeCollection list2 = new RtfNodeCollection();

            RtfTreeNode node = new RtfTreeNode(RtfNodeType.Keyword, "b", true, 2);

            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "a", true, 1));
            list1.Add(node);
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "c", true, 3));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "d", true, 4));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "e", true, 5));

            list2.Add(node);
            list2.Add(new RtfTreeNode(RtfNodeType.Keyword, "g", true, 7));

            Assert.That(list1.Count, Is.EqualTo(5));
            Assert.That(list2.Count, Is.EqualTo(2));

            Assert.That(list1[1], Is.SameAs(node));
            Assert.That(list2[0], Is.SameAs(node));

            list1.AddRange(list2);

            Assert.That(list1.Count, Is.EqualTo(7));

            Assert.That(list1[5], Is.SameAs(list2[0]));
            Assert.That(list1[6], Is.SameAs(list2[1]));

            Assert.That(list1[6].NodeKey, Is.EqualTo("g"));
            Assert.That(list2[0], Is.SameAs(node));

            RtfTreeNode node1 = new RtfTreeNode(RtfNodeType.Keyword, "h", false, 8);

            list1.Insert(5, node1);

            Assert.That(list1.Count, Is.EqualTo(8));
            Assert.That(list1[5], Is.SameAs(node1));

            RtfTreeNode node2 = new RtfTreeNode(RtfNodeType.Keyword, "i", false, 9);

            list1[1] = node2;

            Assert.That(list1.Count, Is.EqualTo(8));
            Assert.That(list1[1], Is.SameAs(node2));
        }

        [Test]
        public void RemoveNodesFromCollection()
        {
            RtfNodeCollection list1 = new RtfNodeCollection();

            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "a", true, 1));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "b", true, 2));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "c", true, 3));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "d", true, 4));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "e", true, 5));

            Assert.That(list1.Count, Is.EqualTo(5));
            Assert.That(list1[1].NodeKey, Is.EqualTo("b"));

            list1.RemoveAt(1);

            Assert.That(list1.Count, Is.EqualTo(4));
            Assert.That(list1[0].NodeKey, Is.EqualTo("a"));
            Assert.That(list1[1].NodeKey, Is.EqualTo("c"));
            Assert.That(list1[2].NodeKey, Is.EqualTo("d"));
            Assert.That(list1[3].NodeKey, Is.EqualTo("e"));

            list1.RemoveRange(1, 2);

            Assert.That(list1.Count, Is.EqualTo(2));
            Assert.That(list1[0].NodeKey, Is.EqualTo("a"));
            Assert.That(list1[1].NodeKey, Is.EqualTo("e"));
        }

        [Test]
        public void SearchNodes()
        {
            RtfNodeCollection list1 = new RtfNodeCollection();
            RtfTreeNode node = new RtfTreeNode(RtfNodeType.Keyword, "c", true, 3);

            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "a", true, 1));
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "b", true, 2));
            list1.Add(node);
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "b", true, 4));
            list1.Add(node);
            list1.Add(new RtfTreeNode(RtfNodeType.Keyword, "e", true, 6));

            Assert.That(list1.IndexOf(node), Is.EqualTo(2));
            Assert.That(list1.IndexOf(new RtfTreeNode()), Is.EqualTo(-1));

            Assert.That(list1.IndexOf(node, 0), Is.EqualTo(2));
            Assert.That(list1.IndexOf(node, 2), Is.EqualTo(2));
            Assert.That(list1.IndexOf(node, 3), Is.EqualTo(4));
            Assert.That(list1.IndexOf(node, 5), Is.EqualTo(-1));

            Assert.That(list1.IndexOf("b", 0), Is.EqualTo(1));
            Assert.That(list1.IndexOf("b", 1), Is.EqualTo(1));
            Assert.That(list1.IndexOf("b", 2), Is.EqualTo(3));
            Assert.That(list1.IndexOf("x", 0), Is.EqualTo(-1));

            Assert.That(list1.IndexOf("b"), Is.EqualTo(1));
            Assert.That(list1.IndexOf("x"), Is.EqualTo(-1));
        }
    }
}
