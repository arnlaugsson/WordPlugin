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
        private WordError[] wordErrors;
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
                //TODO: Setja gamla gildið eitthvað.
                //Word.WdUnderline oldUnderlineFormat = sentence[item.getWordNumber()].Font.Underline;
                //Word.WdColor oldColor = sentence[item.getWordNumber()].Font.UnderlineColor;
                wordUnderline[item.getWordNumber()] = sentence[item.getWordNumber()].Font.Underline;
                wordColor[item.getWordNumber()] = sentence[item.getWordNumber()].Font.UnderlineColor;

                sentence[item.getWordNumber()].Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                sentence[item.getWordNumber()].Font.UnderlineColor = Word.WdColor.wdColorGreen;
            }
        }

        public void resetErrors()
        {
            foreach (WordError item in errorList.getErrorList())
            {
                sentence[item.getWordNumber()].Font.Underline = wordUnderline[item.getWordNumber()];
                sentence[item.getWordNumber()].Font.UnderlineColor = wordColor[item.getWordNumber()];
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
