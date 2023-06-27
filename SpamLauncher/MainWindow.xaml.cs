using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Outlook;


namespace SpamLauncher
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Microsoft.Office.Interop.Outlook.Application outlook;
        private MailItem mail;


        public MainWindow()
        {
            InitializeComponent();
            outlook = new Microsoft.Office.Interop.Outlook.Application();
        }

        /// <summary>
        /// Générer des e-mails
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Créer notre objet mail

            mail = outlook.CreateItem(OlItemType.olMailItem);
            //On définit le sujet du mail

            mail.Subject = "Ceci est un mail de test";
            //On définit le ou les destinataires

            mail.To = TextAdresse.Text; //Récupere la valeur du textbox
            //On définit le corps du message
            mail.Body = TextMessage.Text;
            mail.Display(true);

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (mail != null)
                if (!mail.Sent)
                   mail.Send();
                
                    
            
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            int nBmail = 10;
            for (int i =0; i<nBmail; i++)

            {
                mail = outlook.CreateItem(OlItemType.olMailItem);
               

                mail.Subject = "Mail n" + i;
                

                mail.To = TextAdresse.Text; //Récupere la valeur du textbox
                                            //On définit le corps du message
                mail.Body = TextMessage.Text;
                mail.Send();
                MessageBox.Show("Mail n" + i + "envoyé!");

            }
        }
    }
}
