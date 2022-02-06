using System;

namespace powerpoint_for_translation
{
    class oxd{
        static void Main(string[] args)
        {
            Console.Write("1 --> ToWord, 2 --> ToPowerpoint: ");
            int choiceNum = Int32.Parse(Console.ReadLine());
            Console.Write("Name of File: ");
            string nameChoice = Console.ReadLine();

            if(choiceNum == 1){
                ooxmlDocument.pptx xmlPptx = new ooxmlDocument.pptx(nameChoice);
                xmlPptx.buildDocx();
            }else if(choiceNum == 2){
                ooxmlDocument.docx xmlDocx = new ooxmlDocument.docx(nameChoice);
                xmlDocx.buildPptx();
            }else{
                Console.WriteLine("Not a option");
            }
        }
    }
    
}