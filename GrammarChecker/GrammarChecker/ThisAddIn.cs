using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;
using System.IO;

namespace GrammarChecker
{
    public partial class ThisAddIn
    {
        Word.Document Doc;
        //Erum ekki að nota counterinn eins og er.
        //int counter = 1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Put the active document as the working document.
            Doc = this.Application.ActiveDocument;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void insertText()
        {
            //Testing
            //Doc.Paragraphs[1].Range.InsertParagraphBefore();
            //Doc.Paragraphs[1].Range.Text = "Setjum inn línu " + counter++ + ".";
            
            //String textToParse = "Bráðum koma jólin. Eða er það ekki? Jú víst.";

            //Get the last sentance from the document.
            //String textToParse = Doc.Sentences.Last.Text;
            //String textToParse = Doc.Range().Text;
            Word.Selection currentSelection = Application.Selection;
            String textToParse = currentSelection.Text;
            

            // Setup the process with the ProcessStartInfo class.
            ProcessStartInfo start = new ProcessStartInfo();
            //TODO: Athuga afhverju environment stillingar koma ekki inn. (java finnst ekki nema ég gefi fullan path)
            start.FileName = @"C:\Program Files\Java\jre6\bin\javaw.exe"; // Specify exe file.
            //start.Arguments = "-jar c:\\malvinnsla\\Malgrylan\\GrylaGit\\Grylan\\build\\jar\\Gryla.jar \"" + textToParse + "\"";
            start.Arguments = "-jar C:\\malvinnsla\\Malgrylan\\GrylanGit\\Grylan\\build\\jar\\Gryla.jar \"" + textToParse + "\"";
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            
            //Góðan daginn Skúli vinur minn. Úti er mikið myrkur.

            // Start the process.
            using (Process process = Process.Start(start))
            {
                // Read in all the text from the process with the StreamReader.
                using (StreamReader reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd();

                    //Sendi result í wordskjalið hér... á ekki að vera svoleiðis í framtíðinni, en gott á meðan við útbúum errorlistann.
                    //Doc.Paragraphs[1].Range.Text = result;
                    Doc.Sentences.Last.Text += result;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
