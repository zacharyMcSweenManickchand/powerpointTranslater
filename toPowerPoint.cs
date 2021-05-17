using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace powerpoint_for_translation
{
    class ToPowerpoint
    {
        public static List<(List<string>, List<string>, List<List<List<string>>>)> GetData(string path)
        {
            List<(List<string>, List<string>, List<List<List<string>>>)> grid = new List<(List<string>, List<string>, List<List<List<string>>>)>();
            

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, false))
            {
                Document doc = wordDocument.MainDocumentPart.Document;
                IEnumerable<Table> tableIE = doc.Body.Descendants<Table>();
                foreach(var table in tableIE)
                {
                    List<string> content = new List<string>();
                    List<string> notes = new List<string>();
                    List<List<List<string>>> tables = new List<List<List<string>>>();

                    IEnumerable<TableRow> rows = table.Descendants<TableRow>();
                    List<List<string>> nestedRow = new List<List<string>>();
                    Boolean HasTable = false;
                    foreach (var row in rows)
                    {
                        IEnumerable<TableCell> cells = row.Descendants<TableCell>();
                        List<string> nestedCell = new List<string>();
                        int cellsCount = cells.Count();
                        //Console.WriteLine(cellsCount);
                        for(int i = cellsCount/2; i < cellsCount; i++)
                        {
                            
                            string ogText = cells.ElementAt(i - cellsCount/2).InnerText;
                            string text = cells.ElementAt(i).InnerText; 
                            if(cellsCount == 2){
                                //Content and Notes
                                if( ogText.Length > 7 && ogText.Substring(0, 7) == "Notes: ")
                                {
                                    notes.Add(text);
                                }
                                else
                                {
                                    //Console.WriteLine(text);
                                    //Console.WriteLine("Content");
                                    StringBuilder ContentText = new StringBuilder();
                                    IEnumerable<Paragraph> paragraphs = cells.ElementAt(i).Descendants<Paragraph>();
                                    IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Text> texts = cells.ElementAt(i).Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
                                    foreach(var t in texts)
                                    {
                                        
                                        string tt = t.Text;
                                        string cts = ContentText.ToString();
                                        if (cts == ""){
                                            ContentText.Append(tt);
                                        }
                                        else if(cts[cts.Length - 1] != ' ' && tt[0] != ' '){
                                            ContentText.Append("\n" + tt);
                                        }
                                        else{
                                            ContentText.Append(tt);
                                        }
                                    }
                                    //Console.WriteLine(ContentText.ToString());
                                    content.Add(ContentText.ToString());
                                }
                            }
                            else
                            {
                                //Table
                                HasTable = true;
                                nestedCell.Add(text);
                            }
                            
                        }
                        if(HasTable)
                        {
                            nestedRow.Add(nestedCell);
                        }
                    }
                    if(HasTable)
                    {
                        tables.Add(nestedRow);
                    }

                    grid.Add((content, notes, tables));
                }
                wordDocument.Save();
            }

            return grid;
        }

        public static void editPowerpoint(List<(List<string>, List<string>, List<List<List<string>>>)> grid, string pDoc, string newPDoc)
        {
            File.Copy(pDoc, newPDoc);
            using(PresentationDocument ppt = PresentationDocument.Open(newPDoc, true))
            {
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                for(int i = 0; i < grid.Count(); i++)
                {
                    string relId = (slideIds[i] as SlideId).RelationshipId;
                    SlidePart slide = (SlidePart) part.GetPartById(relId);
                    if(grid.ElementAt(i).Item1.Count() != 0){
                        //Console.WriteLine(grid.ElementAt(i).Item1.Count());
                        editContent(slide, grid.ElementAt(i).Item1);
                    }
                    if(grid.ElementAt(i).Item2.Count() != 0){
                        //Console.WriteLine(grid.ElementAt(i).Item2.Count());
                        editNotes(slide, grid.ElementAt(i).Item2); 
                    }
                    if(grid.ElementAt(i).Item3.Count() != 0){
                        //Console.WriteLine(grid.ElementAt(i).Item3.Count());
                        editTables(slide, grid.ElementAt(i).Item3);
                    }
                    
                }
            }
        }
        private static void editContent(SlidePart slide, List<string> newContent)
        {
            IEnumerable<TextBody> tb = slide.Slide.Descendants<TextBody>();
            if(tb != null)
            {
                for (int i = 0; i < tb.Count(); i++)
                {
                    TextBody inner = tb.ElementAt(i);
                    changeText(inner, newContent.ElementAt(i));
                }
            }
            
        }
        private static void editNotes(SlidePart slide, List<string> newNotes)
        {
            IEnumerable<NotesSlidePart> notes = slide.GetPartsOfType<NotesSlidePart>();
            if(notes != null)
            {
                IEnumerable<TextBody> notesText = notes.First().NotesSlide.Descendants<TextBody>();
                //Console.WriteLine("Number of Textbox in Notes: " + notesText.Count().ToString());
                /*Console.WriteLine(notesText.ElementAt(0).InnerText);
                Console.WriteLine(notesText.ElementAt(1).InnerText);
                for (int i = 0; i < notesText.Count(); i++)
                {
                    string noteText = notesText.ElementAt(i).InnerText;
                    int pageNumber;
                    if(noteText != "" && !int.TryParse(noteText, out pageNumber))
                    {
                        TextBody inner = notesText.ElementAt(i);
                        changeText(inner, newNotes.ElementAt(i));
                    }
                    
                }*/
                if (newNotes.Count() <= notesText.Count())
                {
                    for(int i = 0; i < newNotes.Count(); i++)
                    {
                        string noteText = notesText.ElementAt(i).InnerText;
                        int pageNumber;
                        if(noteText != "" && !int.TryParse(noteText, out pageNumber))
                        {
                            TextBody inner = notesText.ElementAt(i);
                            changeText(inner, newNotes.ElementAt(i));
                        }
                    }  
                }else{
                    Console.WriteLine("You have changed the original file since you have made the word file");
                }
            }
        }
        private static void editTables(SlidePart slide, List<List<List<string>>> newTables)
        {
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
                    IEnumerable<A.Table> tb = graphicData.Descendants<A.Table>();
                    for(int i = 0; i < tb.Count(); i++)
                    {
                        if(tb.ElementAt(i) != null){
                            List<List<string>> newTable = newTables.ElementAt(i);
                            IEnumerable<A.TableRow> tableRows = tb.ElementAt(i).Descendants<A.TableRow>();
                            for(int row = 0; row < tableRows.Count(); row++)
                            {
                                List<string> rows = newTable.ElementAt(row);
                                IEnumerable<A.TableCell> tableCells = tableRows.ElementAt(row).Descendants<A.TableCell>();
                                for(int cell = 0; cell < tableCells.Count(); cell++)
                                {
                                    string value = rows.ElementAt(cell);
                                    A.TextBody textBody = tableCells.ElementAt(cell).Descendants<A.TextBody>().FirstOrDefault();
                                    tableText(textBody, value);
                                } 
                            }
                        }
                    }
                    
                }
            }
        }
        private static void tableText(A.TextBody tb, string message)
        {
            var texts = tb.Descendants<A.Text>();
            if(texts.Count() > 0)
            {
                texts.ElementAt(0).Text = message;
                foreach(var text in texts){
                    if(texts.ElementAt(0) != text){
                        text.Remove();
                    }
                } 
            }
        }
        private static void changeText(TextBody tb, string message)
        {
            var texts = tb.Descendants<A.Text>();
            if(texts.Count() > 0)
            {
                texts.ElementAt(0).Text = message;
                foreach(var text in texts){
                    if(texts.ElementAt(0) != text){
                        text.Remove();
                    }
                } 
            }
            
        }
    }
}