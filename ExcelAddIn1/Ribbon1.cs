using Microsoft.Office.Tools.Ribbon;
using WpfApplication1;


namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MainWindow wpf = new MainWindow();
            wpf.InitializeComponent();
            wpf.Run();
        }

        public void LogPagesOperations(string pageAction)
        {
            this.lblLastDeletedPage.Label = pageAction;
        }
    }
}
