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
 * Class:		RtfDocumentTest
 * Description:	Proyecto de Test para NRtfTree
 * ******************************************************************************/

using System;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;
using System.IO;
using NUnit.Framework;
using System.Drawing;
using System.Globalization;

namespace Net.Sgoliver.NRtfTree.Test
{
    [TestFixture]
    public class RtfDocumentTest
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
        public void CreateSimpleDocument()
        {
            RtfDocument doc = new RtfDocument("..\\..\\testdocs\\rtfdocument.rtf");

            RtfCharFormat charFormat = new RtfCharFormat();
            charFormat.Color = Color.DarkBlue;
            charFormat.Underline = true;
            charFormat.Bold = true;
            doc.UpdateCharFormat(charFormat);

            RtfParFormat parFormat = new RtfParFormat();
            parFormat.Alignment = TextAlignment.Justified;
            doc.UpdateParFormat(parFormat);

            doc.AddText("First Paragraph");
            doc.AddNewParagraph(2);

            doc.SetFormatBold(false);
            doc.SetFormatUnderline(false);
            doc.SetFormatColor(Color.Red);

            doc.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer quis eros at tortor pharetra laoreet. Donec tortor diam, imperdiet ut porta quis, congue eu justo.");
            doc.AddText("Quisque viverra tellus id mauris tincidunt luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.");
            doc.AddText("Donec ac leo justo, vitae rutrum elit. Nulla tellus elit, imperdiet luctus porta vel, consectetur quis turpis. Nam purus odio, dictum vitae sollicitudin nec, tempor eget mi.");
            doc.AddText("Etiam vitae porttitor enim. Aenean molestie facilisis magna, quis tincidunt leo placerat in. Maecenas malesuada eleifend nunc vitae cursus.");
            doc.AddNewParagraph(2);

            doc.AddText("Second Paragraph", charFormat);
            doc.AddNewParagraph(2);

            charFormat.Font = "Courier New";
            charFormat.Color = Color.Green;
            charFormat.Bold = false;
            charFormat.Underline = false;
            doc.UpdateCharFormat(charFormat);

            doc.SetAlignment(TextAlignment.Left);
            doc.SetLeftIndentation(2);
            doc.SetRightIndentation(2);

            doc.AddText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer quis eros at tortor pharetra laoreet. Donec tortor diam, imperdiet ut porta quis, congue eu justo.");
            doc.AddText("Quisque viverra tellus id mauris tincidunt luctus. Fusce in interdum ipsum. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.");
            doc.AddText("Donec ac leo justo, vitae rutrum elit. Nulla tellus elit, imperdiet luctus porta vel, consectetur quis turpis. Nam purus odio, dictum vitae sollicitudin nec, tempor eget mi.");
            doc.AddText("Etiam vitae porttitor enim. Aenean molestie facilisis magna, quis tincidunt leo placerat in. Maecenas malesuada eleifend nunc vitae cursus.");
            doc.AddNewParagraph(2);

            doc.UpdateCharFormat(charFormat);
            doc.SetFormatUnderline(false);
            doc.SetFormatItalic(true);
            doc.SetFormatColor(Color.DarkBlue);

            doc.SetLeftIndentation(0);

            doc.AddText("Test Doc.");
            doc.AddNewLine(2);
            doc.AddText("Stop.");

            doc.Close();

            StreamReader sr = null;
            sr = new StreamReader("..\\..\\testdocs\\rtfdocument.rtf");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("..\\..\\testdocs\\rtf4.txt");
            string rtf4 = sr.ReadToEnd();
            sr.Close();

            //Se adapta el lenguaje al del PC donde se ejecutan los tests
            int deflangInd = rtf4.IndexOf("\\deflang3082");
            rtf4 = rtf4.Substring(0, deflangInd) + "\\deflang" + CultureInfo.CurrentCulture.LCID + rtf4.Substring(deflangInd + 8 + CultureInfo.CurrentCulture.LCID.ToString().Length);

            Assert.That(rtf1, Is.EqualTo(rtf4));
        }
    }
}
