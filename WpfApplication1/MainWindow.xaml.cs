using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public void Run()
        {
            ExcelInstCollector eic = new ExcelInstCollector();
            var bl = new BindingList<string>(eic.ActiveInstances);
            var source = new BindingSource(bl, null);
            this.gridExcelInstances.ItemsSource = source;
            this.Show();
        }
    }
}
