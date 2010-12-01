using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace GrammarChecker
{
    /* This class contains all words for a sentance and can say if it contains errors.
     * If there are errors then there should be a suggestion for how to fix it. */
    class WordCollection
    {
        private Word.Words sentence;
        private ErrorList errorList;
        private int sentenceNumber;
        private Word.WdUnderline[] wordUnderline;
        private Word.WdColor[] wordColor;

        public WordCollection(Word.Words sentence, ErrorList errors, int sentenceNumber)
        {
            this.sentence = sentence;
            this.errorList = errors;
            this.sentenceNumber = sentenceNumber;
            this.wordUnderline = new Word.WdUnderline[sentence.Count];
            this.wordColor = new Word.WdColor[sentence.Count];

            markWordsIfError();
        }

        public void markWordsIfError()
        {
            foreach (WordError item in errorList.getErrorList())
            {
                //Save the current underline state and color.
                wordUnderline[item.getWordNumber()] = sentence[item.getWordNumber()].Font.Underline;
                wordColor[item.getWordNumber()] = sentence[item.getWordNumber()].Font.UnderlineColor;
                //Put curly underline in a green color.
                sentence[item.getWordNumber()].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                sentence[item.getWordNumber()].Font.UnderlineColor = Word.WdColor.wdColorGreen;
            }
        }

        public void resetErrors()
        {
            foreach (WordError item in errorList.getErrorList())
            {
                try
                {
                    sentence[item.getWordNumber()].Font.Underline = wordUnderline[item.getWordNumber()];
                    sentence[item.getWordNumber()].Font.UnderlineColor = wordColor[item.getWordNumber()];
                }
                catch (Exception ex)
                {
                    //Do nothing. Line was probably deleted before the errors where cleared, so we don't care about this line.
                }
            }
        }

        public int getSentenceNumber()
        {
            return this.sentenceNumber;
        }

    }
}
