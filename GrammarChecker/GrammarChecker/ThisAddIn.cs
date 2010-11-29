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
            //Test línur til að parsa
            //Bráðum koma jólin. Eða er það ekki? Jú víst.
            //Góðan daginn Skúli vinur minn. Úti er mikið myrkur.
            //Hún er góði kennarans. Hann er stór strákar. Hún er góð kennari. Hún er góður. Hún hljóp í gegnum skóginum. Hann borðuðu mikið.

            //Check if there is a selected text to parse, if not then we take all document.
            Word.Selection currentSelection = Application.Selection;
            Sentences sentencesToParse;
            if (!currentSelection.Text.Equals("\r"))
            {
                sentencesToParse = currentSelection.Sentences;
            }
            else
            {
                sentencesToParse = Doc.Sentences;
            }

            ArrayList allErrors = new ArrayList();

            //For each sentance we run it through iceparser and keep track of errors.
            for (int sentenceNumber = 1; sentenceNumber < sentencesToParse.Count + 1; sentenceNumber++)
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
                WordCollection sentence = new WordCollection(sentencesToParse[sentenceNumber].Words, errorList, sentenceNumber);
                sentences.Add(sentence);
                foreach (WordError we in errorList.getErrorList())
                {
                    allErrors.Add(we);
                }
            }
            
            //foreach (SentenceContainer sc in sentences)
            //foreach (WordCollection sc in sentences.ToArray())
            //{
            //    //Við ætlum að sækja setningu nr. og athuga hvaða orð eru með villur og setja undirlínu á þau orð.
            //    WordError[] wordErrors = sc.getWordErrors();
            //    if (wordErrors != null)
            //    {
            //        //foreach (WordError item in sc.getWordErrors())
            //        for (int i = 0; i < wordErrors.Length; i++)
            //        {
            //            WordError item = wordErrors[i];
            //            Doc.Sentences[sc.getSentenceNumber()].Words[item.getWordNumber()].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
            //            Doc.Sentences[sc.getSentenceNumber()].Words[item.getWordNumber()].Font.UnderlineColor = Word.WdColor.wdColorGreen;
            //            //Doc.Sentences[sc.getSentenceNumber()].Words[item.getWordNumber()] ;
            //            allErrors.Add(item);
            //        }
                    
            //    }
            //}

            displayErrors(allErrors);
        }

        private void displayErrors(ArrayList allErrors)
        {
            string message = "";
            RuleDescriptions rules = new RuleDescriptions();
            
            foreach (WordError error in allErrors){
                message += "\"" + error.getWord() + "\" violates " + rules.getRule(error.getRuleNumber()) + "\n\n";
            }

            System.Windows.Forms.MessageBox.Show("List of errors:\n" + message);
        }

        /**
         *Function that parses the result string into a WordError object. 
         **/
        private WordError parseResult(String result)
        {
            string[] tokens = result.Split(new char[] { ' ' }, 4);
            tokens[3] = tokens[3].Replace("[", "");
            tokens[3] = tokens[3].Replace("]", "");
            tokens[3] = tokens[3].Replace("\r\n", "");
            string[] suggestions;
            if (!tokens[3].Equals(""))
            {
                suggestions = tokens[3].Split(new char[] { ' ' });
            }
            else
            {
                suggestions = null;
            }

            WordError wordError = new WordError(Convert.ToInt32(tokens[0]) + 1, tokens[1], Convert.ToInt32(tokens[2]), suggestions);
            return wordError;
        }

        private void button1_ClickCheckSelectedText(object sender, Word.Document doc, Word.Window Wn)
        {
            System.Windows.Forms.MessageBox.Show("POPUP GLUGGI");
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
