using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace GrammarChecker
{
    /* This class contains all words for a sentance and can say if it contains errors.
     * If there are errors then there should be a suggestion for how to fix it. */
    class SentenceContainer
    {
        private Word.Words sentence;
        private bool[] markErrors;
        private WordError[] wordErrors;
        private ErrorList errorList;
        private int sentenceNumber;

        public SentenceContainer(Word.Words sentence, ErrorList errors, int sentenceNumber)
        {
            this.sentence = sentence;
            this.errorList = errors;
            this.sentenceNumber = sentenceNumber;

            markErrors = new bool[sentence.Count];
            markWordsIfError();
        }

        public void markWordsIfError()
        {
            foreach (WordError item in errorList.getErrorList())
            {
                sentence[item.getWordNumber()].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                sentence[item.getWordNumber()].Font.UnderlineColor = Word.WdColor.wdColorGreen;
            }
        }

        public int getSentenceNumber()
        {
            return this.sentenceNumber;
        }

        public WordError[] getWordErrors()
        {
            return this.wordErrors;
        }

    }
}
