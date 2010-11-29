using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GrammarChecker
{
    class RuleDescriptions
    {
        public String getRule(int number)
        {
            switch (number)
            {
                case 1: return "Rule 1: Disagreement in case within a noun phrase";
                case 2: return "Rule 2: Disagreement in number within a noun phrase";
                case 3: return "Rule 3: Disagreement in gender within a noun phrase";
                case 4: return "Rule 4 or 5: Disagreement in gender or number between the subject and the complement";
                case 5: return "Rule 4 or 5: Disagreement in gender or number between the subject and the complement";
                case 6: return "Rule 6: Disagreement in case in a prepositional phrase";
                case 7: return "Rule 7: Disagreement in number between subject and related word";
                default: return "Undocumented error";
            }

        }
    }
}
