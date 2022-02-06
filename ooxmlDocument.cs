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
    class ooxmlDocument
    {
        public class pptx 
        {
            public List<slide> slides;
            public string NameOfFile;
            public class slide{
                // change all values to private
                private PresentationPart part;
                public List<List<List<string>>> tables = new List<List<List<string>>>();
                //public List<IEnumerable<A.Table>> tables = new List<IEnumerable<A.Table>>();
                public List<string> content = new List<string>();
                public List<string> notes = new List<string>();
                public int index;
                public slide(PresentationPart p, int i){
                    index = i;
                    part = p;
                    tables = GetTables();
                    content = contentText();
                    notes = getNotes();
                }
                private List<string> contentText()
                {
                    OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                    string relId = (slideIds[index] as SlideId).RelationshipId;

                    SlidePart slide = (SlidePart) part.GetPartById(relId);
                    /*IEnumerable<A.HyperlinkType> links = slide.Slide.Descendants<A.HyperlinkType>(); // Finding links Start
                    if(links.Count() > 0){
                        foreach(var link in links){
                            foreach (HyperlinkRelationship relation in slide.HyperlinkRelationships)
                            {
                                if (relation.Id.Equals(link.Id))
                                {
                                    // Add the URI of the external relationship to the list of strings.
                                    Console.WriteLine(relation.Uri.AbsoluteUri);
                                    Console.WriteLine(link.Id);
                                }
                            }
                        }
                    } // Finding Links end*/

                    IEnumerable<TextBody> tb = slide.Slide.Descendants<TextBody>();
                    List<string> tempArr = new List<string>();
                    if(tb != null)
                    {
                        foreach (TextBody inner in tb)
                        {
                            //Console.WriteLine(inner.InnerText.ToString());
                            StringBuilder paragraphText = new StringBuilder();
                            IEnumerable<A.Paragraph> paragraphs = inner.Descendants<A.Paragraph>();
                            /*foreach(var ol in paragraphs){
                                Console.WriteLine(ol.OuterXml);
                            }*/
                            foreach(A.Paragraph p in paragraphs){
                                //Console.WriteLine(p.OuterXml);
                                IEnumerable<A.Run> runs = p.Descendants<A.Run>();
                                foreach(A.Run run in runs){
                                    //Console.WriteLine(run.OuterXml);
                                    IEnumerable<A.Text> texts = run.Descendants<A.Text>();
                                    IEnumerable<A.HyperlinkOnClick> hrefs = run.Descendants<A.HyperlinkOnClick>();
                                    if(texts.Count() != 0){
                                        foreach(A.Text text in texts){
                                            string pts = paragraphText.ToString();
                                            string tt = text.Text;
                                            //Console.WriteLine(text.OuterXml);
                                            //Links will look like this: [Google](https://www.google.com)
                                            if (hrefs.Count() != 0){
                                                string hrf = String.Format("[{0}]({1})", tt, hrefs.First().Tooltip.ToString());
                                                paragraphText.Append(hrf);
                                            }
                                            else if (pts.ToString() == ""){
                                                paragraphText.Append(tt);
                                            }
                                            else if(pts[pts.Length - 1] != ' ' && tt[0] != ' '){
                                                paragraphText.Append("\n" + tt);
                                            }
                                            else{
                                                paragraphText.Append(tt);
                                            }
                                        }
                                    }
                                }
                                
                            }
                            tempArr.Add(paragraphText.ToString());
                            /*IEnumerable<A.Text> texts = inner.Descendants<A.Text>();
                            foreach (A.Text text in texts)
                            {
                                string pts = paragraphText.ToString();
                                string tt = text.Text;
                                Console.WriteLine(text.OuterXml);
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
                            Console.WriteLine(paragraphText.ToString());//Takeout*/
                        }
                    }
                        
                    return tempArr;       
                }
                private List<string> getNotes()
                {
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
                                //Console.WriteLine("Notes: " + noteText.InnerText);//Takeout
                            }
                        }

                        return note;
                    }

                    //Console.WriteLine("Number of Notes: " + notes.Count().ToString());
                    return note;
                }
                private List<List<List<string>>> GetTables()
                {
                    OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                    string relId = (slideIds[index] as SlideId).RelationshipId;

                    SlidePart slide = (SlidePart) part.GetPartById(relId);

                    List<List<List<string>>> tabs = new List<List<List<string>>>();
                    //List<IEnumerable<A.Table>> tabs = new List<IEnumerable<A.Table>>();
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
                                    List<List<string>> value = new List<List<string>>();
                                    IEnumerable<A.TableRow> tableRows = tb.Descendants<A.TableRow>();
                                    foreach(var tableRow in tableRows)
                                    {
                                        IEnumerable<A.TableCell> tableCells = tableRow.Descendants<A.TableCell>();
                                        List<string> TempList = new List<string>();
                                        foreach(var tableCell in tableCells)
                                        {
                                            A.TextBody textBody = tableCell.Descendants<A.TextBody>().FirstOrDefault();
                                            TempList.Add(textBody.InnerText);
                                            //Console.WriteLine(textBody.InnerText);//Takeout
                                        }
                                        value.Add(TempList);     
                                    }
                                    tabs.Add(value);
                                }
                            }
                            
                        }
                        return tabs;
                    }
                    return tabs;
                }
            }
            public pptx(string filename){ //Constructer
                if(filename != null){        
                    NameOfFile = filename;
                    filename += ".pptx";
                    using (PresentationDocument presentationDocument = PresentationDocument.Open(filename, false))
                    {
                        // Get the presentation part of document.
                        PresentationPart part = presentationDocument.PresentationPart;
                        if (part != null){
                            int slidesCount = part.SlideParts.Count();
                            slides = new List<slide>();
                            for(int i = 0; i < slidesCount; i++){
                                slides.Add(new slide(part, i));
                            } 
                        }else{
                            Console.WriteLine("Presentation Part is equal to Null");
                        }      
                    }
                }else{
                    Console.WriteLine("File name is Null!");
                }
            }
            //Word
            public void buildDocx(){
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(NameOfFile + ".docx", WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    for(int i = 0; i < slides.Count(); i++)
                    {
                        AddTable(wordDocument, slides.ElementAt(i), i + 1);
                    }
                    wordDocument.Save();
                }
            }
            public void AddTable(WordprocessingDocument wordDocument, slide data, int index)
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
                    //new TableCellSpacing() { Width = "100%"},
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
                //100% of the page
                //TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                table.AppendChild<TableProperties>(props);
                //table.AppendChild<TableWidth>(tableWidth);
                //string outerP = "<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr algn=\"ctr\"><a:lnSpc><a:spcPct val=\"100000\" /></a:lnSpc><a:defRPr /></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"3200\" b=\"0\" strike=\"noStrike\" spc=\"-1\"><a:latin typeface=\"Arial\" /></a:rPr><a:t>MyFirstContent</a:t></a:r><a:endParaRPr lang=\"en-US\" sz=\"3200\" b=\"0\" strike=\"noStrike\" spc=\"-1\"><a:latin typeface=\"Arial\" /></a:endParaRPr></a:p>";
                //table.Append(new TableRow(new TableCell(new Paragraph(outerP))));
                foreach (var i in data.content)//Content
                {
                    var tr = new TableRow();
                    var tc = new TableCell();
                    var blank = new TableCell();
                    tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(i))));
                    blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                    blank.Append(new TableCellProperties(
                        new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = "50%" }));
                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "50%"}));
                    tr.Append(tc);
                    tr.Append(blank);
                    table.Append(tr);
                }
                foreach(var i in data.notes)//Notes
                {
                    var tr = new TableRow();
                    var tc = new TableCell();
                    var blank = new TableCell();
                    tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(i))));
                    blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                    // Assume you want columns that are automatically sized.
                    blank.Append(new TableCellProperties(
                        new TableCellWidth() { Type = TableWidthUnitValues.Auto, Width = "50%" }));
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "50%" }));
                    
                    tr.Append(tc);
                    tr.Append(blank);
                    table.Append(tr);
                }
                foreach(List<List<string>> tbls in data.tables) //Tables
                {
                    int rowCount = tbls.Count();
                    TableRow tr = new TableRow(); //Top Row
                    Table newTbl = new Table();
                    foreach(var row in tbls)
                    {   // this makes the the cell to the number of a cells in the row into a percentage | ex: if you have 3 cell then it will output 33%
                        string cellFraction = ((1.0/(double)(row.Count()/**2*/))*100).ToString() + "%"; 
                        TableRow lr = new TableRow(); //Local Row
                        foreach(var cell in row)
                        {
                            var tc = new TableCell();
                            tc.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cell))));
                            //Console.WriteLine(cell);
                            tc.Append(new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = cellFraction }));
                            lr.Append(tc);
                        }
                        newTbl.Append(lr);
                    }
                    var Tcell = new TableCell(); //TopCell
                    Tcell.Append(newTbl);
                    tr.Append(Tcell);

                    var BlankTCell = new TableCell();
                    var BlankT = new Table();
                    foreach(var r in tbls){
                        string cellFraction = ((1.0/(double)(r.Count()/**2*/))*100).ToString() + "%";
                        TableRow lrBlank = new TableRow(); //Local Row
                        for(int i = 0; i < r.Count(); i++){
                            var blank = new TableCell();
                            blank.Append(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(""))));
                            blank.Append(new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = cellFraction }));
                            lrBlank.Append(blank);
                        }
                        BlankT.Append(lrBlank);
                    }
                    BlankTCell.Append(BlankT);
                    tr.Append(BlankTCell);

                    table.Append(tr);
                }

                doc.Body.Append(table);
                doc.Save();
            }
        }
        public class docx
        {
            public List<slideTable> slides;
            public string NameOfFile;
            
            public class slideTable
            {
                public List<List<List<string>>> tables = new List<List<List<string>>>();
                public List<string> content = new List<string>();
                public List<string> notes = new List<string>();
                public Table tbl;
                public slideTable(Table t)
                {
                    tbl = t;
                    IEnumerable<TableRow> rows = tbl.Descendants<TableRow>();
                    foreach(TableRow row in rows)
                    {
                        TableCell cell = row.Descendants<TableCell>().ElementAt(1);//Assuming there are 2 cells in the row this will select the 2nd cell
                        var ce = row.ChildElements.FirstOrDefault();//It gives a IEnumerable with 3 identical result, thus this will pick the First One

                        //Console.WriteLine(e.Parent.Parent.Parent);// This checks if it is in a Nested Table if this outputs TableCell then it is Nested, but if the ouput is Body, then it is not nested
                        //IEnumerable<Table> nestedTable = cell.Descendants<Table>();
                        if(ce.Parent.Parent.Parent.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Body"){// this will ignore the TableCells in Nested Tables
                            IEnumerable<Table> nestedTable = row.Descendants<Table>();
                            if(nestedTable.Count() > 0){
                                Table nTable = nestedTable.ElementAt(1); // Picks the 2nd cell (User input cell), this only works if their is only 2 cells/tables
                                tables.Add(GetNestedTable(nTable));
                            }else{
                                string cellString = cell.InnerText;
                                //Console.WriteLine(cellString);
                                if (cellString.Count() > 7 && cellString.Substring(0, 7).Equals("Notes: ")){
                                    notes.Add(cellString);
                                }else{
                                    content.Add(cellString);
                                }
                            }
                        }
                        
                    }
                }
                private List<List<string>> GetNestedTable(Table nestedTable){
                    List<List<string>> outTable = new List<List<string>>();
                    IEnumerable<TableRow> tblRows = nestedTable.Descendants<TableRow>();
                    foreach(TableRow tblRow in tblRows){
                        List<string> outRow = new List<string>();
                        IEnumerable<TableCell> tblCells = tblRow.Descendants<TableCell>();
                        foreach(TableCell tblCell in tblCells){
                            //Console.WriteLine(tblCell.InnerText);
                            outRow.Add(tblCell.InnerText);
                        }
                        outTable.Add(outRow);
                    }
                    return outTable;
                }
            }
            public docx(string filename)
            {
                if(filename != null){        
                    NameOfFile = filename;
                    filename += ".docx";
                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, false))
                    {
                        // Get the presentation part of document.
                        Document doc = wordDocument.MainDocumentPart.Document;
                        if (doc != null){
                            IEnumerable<Table> tableIE = doc.Body.Descendants<Table>();
                            slides = new List<slideTable>();
                            for(int i = 0; i < tableIE.Count(); i++){
                                //Console.WriteLine(tableIE.ElementAt(i).Parent.ToString() == "DocumentFormat.OpenXml.Wordprocessing.TableCell");
                                if(tableIE.ElementAt(i).Parent.ToString() != "DocumentFormat.OpenXml.Wordprocessing.TableCell"){
                                    slides.Add(new slideTable(tableIE.ElementAt(i)));
                                }/*else{
                                    Console.WriteLine(tableIE.ElementAt(i));
                                }*/
                            }
                        }else{
                            Console.WriteLine("Presentation Part is equal to Null");
                        }      
                    }
                }else{
                    Console.WriteLine("File name is Null!");
                }
            }
            public void buildPptx(){
                if(NameOfFile != null){        
                    string originalPptx = NameOfFile + ".pptx";
                    string newPptx = NameOfFile + "_fr.pptx"; //Change to selected Languafue in the future
                    File.Copy(originalPptx, newPptx);
                    using (PresentationDocument presentationDocument = PresentationDocument.Open(newPptx, true))
                    {
                        // Get the presentation part of document.
                        PresentationPart part = presentationDocument.PresentationPart;
                        OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                        for(int i = 0; i < slides.Count(); i++){
                            string relId = (slideIds[i] as SlideId).RelationshipId;
                            SlidePart slide = (SlidePart) part.GetPartById(relId);
                            if(slides.ElementAt(i).content.Count() > 0){
                                editContent(slide, slides.ElementAt(i).content);
                            }
                            if(slides.ElementAt(i).notes.Count() > 0){
                                editNotes(slide, slides.ElementAt(i).notes);
                            }
                            if(slides.ElementAt(i).tables.Count() > 0){
                                editTables(slide, slides.ElementAt(i).tables);
                            }
                        }  
                    }
                }else{
                    Console.WriteLine("File name is Null!");
                }
            }

            private void editContent(SlidePart slide, List<string> newContent)
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
            private void editNotes(SlidePart slide, List<string> newNotes)
            {
                IEnumerable<NotesSlidePart> notes = slide.GetPartsOfType<NotesSlidePart>();
                if(notes != null)
                {
                    IEnumerable<TextBody> notesText = notes.First().NotesSlide.Descendants<TextBody>();
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
                        Console.WriteLine("You have changed the original powerpoint file since you have made the word file");
                    }
                }
            }
            private void editTables(SlidePart slide, List<List<List<string>>> newTables)
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
            private void tableText(A.TextBody tb, string message)
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
            private void changeText(TextBody tb, string message)
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
}