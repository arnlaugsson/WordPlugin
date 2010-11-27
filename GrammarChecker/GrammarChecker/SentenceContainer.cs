using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace GrammarChecker
{
    class SentenceContainer
    {
        private Word.Words sentence;
        private bool[] markError;
        private WordError[] error;

        public SentenceContainer(Word.Words sentence, ErrorList errors)
        {
            this.sentence = sentence;

            markError = new bool[sentence.Count];
        }
        

        public void fall()
        {
            String ord = sentence.ToString();
            int i = 0;
        }

    }
}
