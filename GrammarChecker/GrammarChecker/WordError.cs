using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GrammarChecker
{
    class WordError
    {
        private int ruleNumber;
        private string[] corrections;

        public WordError(int ruleNumber, string[] corrections)
        {
            this.ruleNumber = ruleNumber;
            this.corrections = corrections;
        }

        public int getRuleNumber() 
        {
            return ruleNumber;
        }

        public string[] getcorrections()
        {
            return corrections;
        }
    }
}
