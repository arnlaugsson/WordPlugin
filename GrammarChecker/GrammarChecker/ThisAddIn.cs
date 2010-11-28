using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace GrammarChecker
{
    public partial class ThisAddIn
    {
        Word.Document Doc;
        ArrayList sentences = new ArrayList();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Put the active document as the working Doc.
            Doc = this.Application.ActiveDocument;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //Bye, bye ...
        }

        public void insertText()
        {
            //Testing
            //Doc.Paragraphs[1].Range.InsertParagraphBefore();
            //Doc.Paragraphs[1].Range.Text = "Setjum inn línu " + counter++ + ".";

            //String textToParse = "Bráðum koma jólin. Eða er það ekki? Jú víst.";
            //Góðan daginn Skúli vinur minn. Úti er mikið myrkur.

            //Check if there is a selected text to parse, if not then we take all document.
            Word.Selection currentSelection = Application.Selection;
            Sentences sentencesToParse;
            if (!currentSelection.Equals(""))
            {
                sentencesToParse = currentSelection.Sentences;
            }
            else
            {
                sentencesToParse = Doc.Sentences;
            }


            //For each sentance we run it through iceparser and keep track of errors.
            for (int sentenceNumber = 1; sentenceNumber < sentencesToParse.Count+1; sentenceNumber++) 
            {
                string textToParse = sentencesToParse[sentenceNumber].Text;
                // Setup the process with the ProcessStartInfo class.
                ProcessStartInfo start = new ProcessStartInfo();
                //TODO: Athuga afhverju environment stillingar koma ekki inn. (java finnst ekki nema ég gefi fullan path)
                start.FileName = @"C:\Program Files\Java\jre6\bin\javaw.exe"; // Specify exe file.
                start.Arguments = "-jar C:\\malvinnsla\\Malgrylan\\GrylanGit\\Grylan\\build\\jar\\Gryla.jar \"" + textToParse + "\"";
                start.UseShellExecute = false;
                start.RedirectStandardOutput = true;

                //Variable that gets the result. 
                string result = "";
                // Start the process.
                using (Process process = Process.Start(start))
                {
                    // Read in all the text from the process with the StreamReader.
                    using (StreamReader reader = process.StandardOutput)
                    {
                        result = reader.ReadToEnd();
                        ////Print out the error list. (for debuging purpose).
                        //System.Windows.Forms.MessageBox.Show("Villulisti:\n" + result);
                    }
                } 

                //ErrorList collects all WordErrors with its parameters (number of word, the word, rulenumber and suggestions)
                ErrorList errorList = new ErrorList();
                if (!result.Equals("") && !result.StartsWith("ok"))
                {
                    //For each line in result we create an error and put it on the ErrorList.
                    errorList.add(parseResult(result));
                }

                //Create the sentence class instance, Values are: an array of words, List of errors, number of the sentance.
                SentenceContainer sentence = new SentenceContainer(sentencesToParse[sentenceNumber].Words, errorList, sentenceNumber);
                sentences.Add(sentence);
            }

            foreach (SentenceContainer sc in sentences)
            {
                //Við ætlum að sækja setningu nr. og athuga hvaða orð eru með villur og setja undirlínu á þau orð.
                foreach (WordError item in sc.getWordErrors())
                {
                    //sentence[item.getWordNumber()].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                    //sentence[item.getWordNumber()].Font.UnderlineColor = Word.WdColor.wdColorGreen;
                }
                //Doc.Sentences[sc.getSentenceNumber()].Words[
            }
        }

        /**
         *Function that parses the result string into a WordError object. 
         **/
        private WordError parseResult(String result)
        {
            string[] tokens = result.Split(new char[] { ' ' }, 4);
            tokens[4] = tokens[4].Replace("[", "");
            tokens[4] = tokens[4].Replace("]", "");
            string[] suggestions = tokens[4].Split(new char[] { ' ' });
            WordError wordError = new WordError(Convert.ToInt32(tokens[0]), tokens[1], Convert.ToInt32(tokens[3]), suggestions);
            return wordError;
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
