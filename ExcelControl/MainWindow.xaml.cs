using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;


namespace ExcelControl
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook classeur;
        private Worksheet feuille;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Le bouton qui lance Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
        }

        /// <summary>
        /// Ferme Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (excel != null)
            {
                excel.Quit();
            }
        }

        /// <summary>
        /// Méthode de traitement
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (excel != null)
            {
                classeur = excel.Workbooks.Add();
                feuille = classeur.ActiveSheet;
                feuille.Cells[2, 2] = "Hello world!";
            }
        }

        /// <summary>
        /// Bouton Couleur
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (feuille != null)
            {
               
                feuille.Cells[2, 2].Interior.Color = XlRgbColor.rgbDarkGreen;

            }
                
        }

        /// <summary>
        /// Créer une table de multiplication
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Button_Click(sender, e);
            Button_Click_2(sender, e);
            int max = 3;
            if (feuille != null) 
            {
                for (int l = 1; l <= max; l++)
                {

                    for (int c = 1; c <= max; c++)
                    {
                        feuille.Cells[l, c] = l * c;
                       
                    }
                }
                Range lignes = feuille.Range[feuille.Cells[1, 1], feuille.Cells[max, max]];
                lignes.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
                lignes.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                lignes.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDouble;
                lignes.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;
                lignes.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDouble;
            }

        }
        /// <summary>
        /// Ferme l'application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
