using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace powerpoint_for_translation
{
    public class toWord
    {
        public static int CountSlides(string presentationFile)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                if (presentationDocument == null)
                {
                    throw new ArgumentNullException("presentationDocument");
                }

                int slidesCount = 0;

                // Get the presentation part of document.
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                // Get the slide count from the SlideParts.
                if (presentationPart != null)
                {
                    slidesCount = presentationPart.SlideParts.Count();
                }
                // Return the slide count to the previous method.
                return slidesCount;
            }
        }
        public static List<string> contentText(string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                SlidePart slide = (SlidePart) part.GetPartById(relId);

                IEnumerable<TextBody> tb = slide.Slide.Descendants<TextBody>();
                List<string> tempArr = new List<string>();
                if(tb != null)
                {
                    foreach (var inner in tb)
                    {
                        //Console.WriteLine(inner.InnerText.ToString());
                        StringBuilder paragraphText = new StringBuilder();
                        IEnumerable<A.Text> texts = inner.Descendants<A.Text>();
                        foreach (A.Text text in texts)
                        {
                            string pts = paragraphText.ToString();
                            string tt = text.Text;
                            if (pts.ToString() == ""){
                                paragraphText.Append(tt);
                            }
                            else if(pts[pts.Length - 1] != ' ' && tt[0] != ' '){
                                paragraphText.Append("\n" + tt);
                            }
                            else{
                                paragraphText.Append(tt);
                            }

                        }
                        tempArr.Add(paragraphText.ToString());
                    }
                }
                
                return tempArr;
            }        
        }
        public static List<string> GetNotes(string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                SlidePart slide = (SlidePart) part.GetPartById(relId);

                List<string> note = new List<string>();
                IEnumerable<NotesSlidePart> notes = slide.GetPartsOfType<NotesSlidePart>();
                if(notes.Count() != 0)
                {
                    IEnumerable<TextBody> notesText = notes.First().NotesSlide.Descendants<TextBody>();
                    //Console.WriteLine("Number of Textbox in Notes: " + notesText.Count().ToString());
                    foreach(var noteText in notesText)
                    {
                        //Console.WriteLine(noteText.InnerText.ToString());
                        int pageNumber;
                        if(noteText.InnerText != "" && !int.TryParse(noteText.InnerText, out pageNumber))
                        {
                            note.Add("Notes: " + noteText.InnerText);
                        }
                    }

                    return note;
                }

                //Console.WriteLine("Number of Notes: " + notes.Count().ToString());
                return note;
            }
        }
        public static List<List<List<string>>> GetTables(string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                SlidePart slide = (SlidePart) part.GetPartById(relId);

                List<List<List<string>>> tables = new List<List<List<string>>>();
                DocumentFormat.OpenXml.Presentation.CommonSlideData commonslideData = slide.Slide.Descendants<CommonSlideData>().FirstOrDefault();
                ShapeTree shapeTree = commonslideData.Descendants<ShapeTree>().FirstOrDefault();
                IEnumerable<DocumentFormat.OpenXml.Presentation.GraphicFrame> graphicFrame = shapeTree.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>();//.ElementAt(tableIndex);
                if(graphicFrame != null)
                {
                    foreach(var gf in graphicFrame)
                    {
                        A.Graphic graphic = gf.Descendants<A.Graphic>().FirstOrDefault();
                        A.GraphicData graphicData = graphic.Descendants<A.GraphicData>().FirstOrDefault<A.GraphicData>();
                        //Table tb = graphicData.Descendants<Table>().FirstOrDefault();
                        IEnumerable<A.Table> tbs = graphicData.Descendants<A.Table>();
                        foreach(var tb in tbs)
                        {
                            if(tb != null){
                                tables.Add(GetContent(tb));
                            }
                        }
                        
                    }
                    return tables;
                }
                return tables;
            }
        }
        private static List<List<string>> GetContent(A.Table table)
        {
            List<List<string>> value = new List<List<string>>();
            IEnumerable<A.TableRow> tableRows = table.Descendants<A.TableRow>();
            foreach(var tableRow in tableRows)
            {
                IEnumerable<A.TableCell> tableCells = tableRow.Descendants<A.TableCell>();
                List<string> TempList = new List<string>();
                foreach(var tableCell in tableCells)
                {
                    A.TextBody textBody = tableCell.Descendants<A.TextBody>().FirstOrDefault();
                    TempList.Add(textBody.InnerText);
                    
                }
                value.Add(TempList);     
            }
            return value;
        }

        //Word
        public static void editWord(string path,  List<(List<string>, List<string>, List<List<List<string>>>)> grid)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                //Paragraph para = body.AppendChild(new Paragraph());
                //Run run = para.AppendChild(new Run());
                //run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("Create text in body - CreateWordprocessingDocument"));
                for(int i = 1; i <= grid.Count(); i++)
                {
                    AddTable(wordDocument, grid.ElementAt(i - 1), i);
                }
                /*int i = 0;
                foreach(var table in grid)
                {
                    i++;
                    AddTable(wordDocument, table, i);
                    //Console.WriteLine("[{0}]", string.Join(", ", table.Item1));
                }*/
                wordDocument.Save();
            }
        }
        public static void AddTable(WordprocessingDocument wordDocument, (List<string>, List<string>, List<List<List<string>>>) data, int index)
        {
            var doc = wordDocument.MainDocumentPart.Document;
            
            string txt;
            //One liner to Add text Above Table | or to simply Add spacing
            if(index != 1)
            {
                txt = String.Format("\nSlide #{0}", index); 
            }
            else
            {
                txt = String.Format("Slide #{0}", index);
            }
            
            doc.Body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(txt));
            Table table = new Table();

            TableProperties props = new TableProperties(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new LeftBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new RightBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideHorizontalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideVerticalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
            }));

            table.AppendChild<TableProperties>(props);

            foreach (var i in data.Item1)
            {
                var tr = new TableRow();
                var tc = new TableCell();
                var blank = new TableCell();
                tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(i))));
                blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                // Assume you want columns that are automatically sized.
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                
                tr.Append(tc);
                tr.Append(blank);
                table.Append(tr);
            }
            foreach(var i in data.Item2)
            {
                var tr = new TableRow();
                var tc = new TableCell();
                var blank = new TableCell();
                tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(i))));
                blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                // Assume you want columns that are automatically sized.
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                
                tr.Append(tc);
                tr.Append(blank);
                table.Append(tr);
            }
            foreach(var tbl in data.Item3)
            {
                foreach(var row in tbl)
                {
                    var tr = new TableRow();
                    foreach(var cell in row)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cell))));
                        tr.Append(tc);
                    }
                    for(int i = 0; i < row.Count(); i++)
                    {
                        var blank = new TableCell();
                        blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                        tr.Append(blank);
                    }
                    table.Append(tr);
                }
            }

            doc.Body.Append(table);
            doc.Save();
        }
    }
}