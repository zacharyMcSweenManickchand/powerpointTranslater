using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace powerpoint_for_translation
{
    class Program
    {
        public static List<(List<string>, List<string>, List<List<List<string>>>)> gridList = new List<(List<string>, List<string>, List<List<List<string>>>)>();
        static void Main(string[] args)
        {
            Console.Write("1 --> ToWord, 2 --> ToPowerpoint: ");
            int choiceNum = Int32.Parse(Console.ReadLine());
            //int choiceNum = 2;
            
            if(choiceNum == 1)
            {
                Console.Write("Please enter a presentation file name without extension: ");
                string fileName = Console.ReadLine();
                string file = @"C:\Users\Public\Documents\" + fileName + ".pptx";
                int numberOfSlides = toWord.CountSlides(file);
                //Console.WriteLine("Number of slides = {0}", numberOfSlides);

                //string[][] gridArr = new string[numberOfSlides][];
                for (int i = 0; i < numberOfSlides; i++)
                {
                    List<string> content = toWord.contentText(file, i);
                    List<string> notes = toWord.GetNotes(file, i);
                    List<List<List<string>>> tables = toWord.GetTables(file, i);
                    gridList.Add((content, notes, tables));
                    //gridArr[i] = slideArr;
                }

                //Prints text by Slide | For Every Slide --> New Grid in Word
                /*foreach(string[] grid in gridArr)
                {
                    Console.WriteLine("[{0}]", string.Join(", ", grid));
                }*/

                toWord.editWord(@"C:\Users\Public\Documents\" + fileName + ".docx", gridList);
            }
            else if(choiceNum == 2)
            {
                Console.Write("Please enter a wordprocessing file name without extension: ");
                string docxFile = Console.ReadLine();
                Console.Write("Please enter a presentation file name without extension: ");
                string pptxFile = Console.ReadLine();
                Console.Write("Please enter a new presentation file name without extension: ");
                string newPptxFile = Console.ReadLine();

                string wordFile = @"C:\Users\Public\Documents\" + docxFile + ".docx";
                string ppFile = @"C:\Users\Public\Documents\" + pptxFile + ".pptx";
                string newPpFile = @"C:\Users\Public\Documents\" + newPptxFile + ".pptx";

                gridList = ToPowerpoint.GetData(wordFile);

                ToPowerpoint.editPowerpoint(gridList, ppFile, newPpFile);
            }
            else if(choiceNum == 3)
            {
                //Test environment
                Test.init();
            }

        }
    }
}
