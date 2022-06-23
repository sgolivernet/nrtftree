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
        public void CreateSimpleDocument()
        {
            RtfDocument doc = new RtfDocument();

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

            doc.Save("testdocs\\rtfdocument1.rtf");

            string text1 = doc.Text;
            string rtfcode1 = doc.Rtf;

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

            doc.AddText("Test Doc. Петяв ñáéíó\n");
            doc.AddNewLine(1);
            doc.AddText("\tStop.");

            string text2 = doc.Text;
            string rtfcode2 = doc.Rtf;

            doc.Save("testdocs\\rtfdocument2.rtf");

            StreamReader sr = null;
            sr = new StreamReader("testdocs\\rtfdocument1.rtf");
            string rtf1 = sr.ReadToEnd();
            sr.Close();

            sr = null;
            sr = new StreamReader("testdocs\\rtfdocument2.rtf");
            string rtf2 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf4.txt");
            string rtf4 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\rtf6.txt");
            string rtf6 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\doctext1.txt");
            string doctext1 = sr.ReadToEnd();
            sr.Close();

            sr = new StreamReader("testdocs\\doctext2.txt");
            string doctext2 = sr.ReadToEnd() + " Петяв ñáéíó\r\n\r\n\tStop.\r\n";
            sr.Close();

            //Se adapta el lenguaje al del PC donde se ejecutan los tests
            int deflangInd = rtf4.IndexOf("\\deflang3082");
            rtf4 = rtf4.Substring(0, deflangInd) + "\\deflang" + CultureInfo.CurrentCulture.LCID + rtf4.Substring(deflangInd + 8 + CultureInfo.CurrentCulture.LCID.ToString().Length);

            //Se adapta el lenguaje al del PC donde se ejecutan los tests
            int deflangInd2 = rtf6.IndexOf("\\deflang3082");
            rtf6 = rtf6.Substring(0, deflangInd2) + "\\deflang" + CultureInfo.CurrentCulture.LCID + rtf6.Substring(deflangInd2 + 8 + CultureInfo.CurrentCulture.LCID.ToString().Length);

            Assert.That(rtf1, Is.EqualTo(rtf6));
            Assert.That(rtf2, Is.EqualTo(rtf4));

            Assert.That(text1, Is.EqualTo(doctext1));
            Assert.That(text2, Is.EqualTo(doctext2));
        }
    }
}
