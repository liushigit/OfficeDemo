using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Word = Microsoft.Office.Interop.Word;


namespace OfficeDemo
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

        private void makeDoc(object sender, RoutedEventArgs e)
        {
            var wordApp = new Word.Application();
            var report = new Report(templatePath: "C:\\Users\\liush_000\\Desktop\\mydoc3.docx", 
                                    outputPath: "C:\\Users\\liush_000\\Desktop\\mydoc5.docx", 
                                    app: wordApp);
            report.SetBookmarkNamed("name", nameTextBox.Text);
            report.SetBookmarkNamed("age", ageTextBox.Text);
            report.SaveOut();

            ((Word._Application)wordApp).Quit(false);
            wordApp = null;
            GC.Collect();
        }
    }
}
