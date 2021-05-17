using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace powerpoint_for_translation
{
    public class Test
    {
        public static void init()
        {
            Console.WriteLine("Hello");
            //string wordFile = @"C:\Users\Public\Documents\" + docxFile + ".docx";
            string ppFile = @"C:\Users\Public\Documents\Sample_en.pptx";
            string newPpFile = @"C:\Users\Public\Documents\newSample.pptx";
            editPowerpoint("hello", ppFile, newPpFile);
        }
        public static void editPowerpoint(string message, string pDoc, string newPDoc)
        {
            File.Copy(pDoc, newPDoc);
            using(PresentationDocument ppt = PresentationDocument.Open(newPDoc, true))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;
                SlidePart slide = (SlidePart) part.GetPartById(relId);
                IEnumerable<TextBody> tb = slide.Slide.Descendants<TextBody>();
                if(tb != null)
                {
                    Console.WriteLine(tb.Count());
                    for (int i = 0; i < tb.Count(); i++)
                    {
                        TextBody inner = tb.ElementAt(i);

                        string innerText = inner.InnerText;
                        //string modifiedString = inner.InnerText.Replace(innerText, message + i.ToString());
                        var tt = inner.Descendants<A.Text>();
                        tt.ElementAt(0).Text = message + i.ToString();
                        foreach(var hg in tt){
                            if(tt.ElementAt(0) != hg){
                                hg.Remove();
                            }
                        }
                    }
                }
            }
        }
    }
}