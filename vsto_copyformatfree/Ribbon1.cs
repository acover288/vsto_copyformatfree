using Microsoft.Office.Tools.Ribbon;

namespace vsto_copyformatfree
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.click();
        }
    }
}
