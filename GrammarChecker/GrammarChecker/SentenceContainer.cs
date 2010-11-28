using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace GrammarChecker
{
    /* This class contains all words for a centance and can say if it contains errors.
     * If there are errors then there should be a suggestion for how to fix it. */
    class SentenceContainer
    {
        private Word.Words sentence;
        private bool[] markError;
        private WordError[] error;
        private ErrorList errors;

        public SentenceContainer(Word.Words sentence, ErrorList errors)
        {
            this.sentence = sentence;
            this.errors = errors;

            markError = new bool[sentence.Count];
        }
        

        public void fall() //TODO: Henda út eftir testing.
        {
            String ord = sentence.ToString();
            int i = 0;
        }

        public void markWordsIfError()
        {
            //Sauðakóði
            //foreach (error in errors) {
            //    sentence[error.number].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
            //    sentence[error.number].Font.UnderlineColor = Word.WdColor.wdColorGreen;
            //}
        }

    }
}
