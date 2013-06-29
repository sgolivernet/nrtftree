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
 * Class:		RtfPullParserTest
 * Description:	Proyecto de Test para RtfPullParser
 * ******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Net.Sgoliver.NRtfTree.Core;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class RtfPullParserTest
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
        public void ParseSimpleDocument()
        {
            RtfPullParser parser = new RtfPullParser();
            parser.LoadRtfFile("..\\..\\testdocs\\testdoc1.rtf");

            parserTests(parser);
        }

        [Test]
        public void ParseSimpleRtfText()
        {
            RtfTree tree = new RtfTree();
            tree.LoadRtfFile("..\\..\\testdocs\\testdoc1.rtf");

            RtfPullParser parser = new RtfPullParser();
            parser.LoadRtfText(tree.Rtf);

            parserTests(parser);
        }

        private static void parserTests(RtfPullParser parser)
        {
            int eventType = parser.GetEventType();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.START_DOCUMENT));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.START_GROUP));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.KEYWORD));
            Assert.That(parser.GetName(), Is.EqualTo("rtf"));
            Assert.That(parser.HasParam(), Is.EqualTo(true));
            Assert.That(parser.GetParam(), Is.EqualTo(1));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.KEYWORD));
            Assert.That(parser.GetName(), Is.EqualTo("ansi"));
            Assert.That(parser.HasParam(), Is.EqualTo(false));

            for (int i = 0; i < 3; i++)
                parser.Next();

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.START_GROUP));

            for (int i = 0; i < 6; i++)
                parser.Next();

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("Times New Roman;"));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.END_GROUP));

            for (int i = 0; i < 27; i++)
                parser.Next();

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.CONTROL));
            Assert.That(parser.GetName(), Is.EqualTo("*"));
            Assert.That(parser.HasParam(), Is.EqualTo(false));

            for (int i = 0; i < 40; i++)
                parser.Next();

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.CONTROL));
            Assert.That(parser.GetName(), Is.EqualTo("'"));
            Assert.That(parser.HasParam(), Is.EqualTo(true));
            Assert.That(parser.GetParam(), Is.EqualTo(233));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("st1 a"));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.CONTROL));
            Assert.That(parser.GetName(), Is.EqualTo("'"));
            Assert.That(parser.HasParam(), Is.EqualTo(true));
            Assert.That(parser.GetParam(), Is.EqualTo(241));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("u "));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("{"));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("\\"));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("test2"));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.TEXT));
            Assert.That(parser.GetText(), Is.EqualTo("}"));

            for (int i = 0; i < 29; i++)
                parser.Next();

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.END_GROUP));

            eventType = parser.Next();
            Assert.That(eventType, Is.EqualTo(RtfPullParser.END_DOCUMENT));
        }
    }
}
