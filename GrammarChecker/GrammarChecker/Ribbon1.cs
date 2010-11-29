using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

using Microsoft.Office.Tools.Word;

namespace GrammarChecker
{
    public partial class GrammarRibbon
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void button1_ClickCheckSelectedText(object sender, RibbonControlEventArgs e)
        {   
            //We need to use ThisAddIn to be able to use ActiveDocument.
            Globals.ThisAddIn.insertText();
        }

        private void button1ResetErrors_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.resetErrors();
        }

    }
}
