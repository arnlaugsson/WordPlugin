using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GrammarChecker
{
    class WordError
    {
        private int wordNumber;
        private string word;
        private int ruleNumber;
        private string[] corrections;

        public WordError(int wordNumber, string word, int ruleNumber, string[] corrections)
        {
            this.wordNumber = wordNumber;
            this.word = word; 
            this.ruleNumber = ruleNumber;
            this.corrections = corrections;
        }

        public int getWordNumber()
        {
            return this.wordNumber;
        }

        public string getWord()
        {
            return this.word;
        }

        public int getRuleNumber() 
        {
            return this.ruleNumber;
        }

        public string[] getcorrections()
        {
            return this.corrections;
        }
    }
}
