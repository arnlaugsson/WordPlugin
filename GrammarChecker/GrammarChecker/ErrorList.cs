using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

namespace GrammarChecker
{
    class ErrorList
    {
        private ArrayList errorList;

        public ErrorList()
        {
            errorList = new ArrayList();
        }

        public void add(WordError wordError)
        {
            errorList.Add(wordError);
        }

        public ArrayList getErrorList()
        {
            return errorList;
        }
    }
}
