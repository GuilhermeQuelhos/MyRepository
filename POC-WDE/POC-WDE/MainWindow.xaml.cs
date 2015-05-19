using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
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
using SHDocVw;

namespace POC_WDE
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenLink("https://mail.google.com/mail/u/1/?tab=wm&pli=1#inbox");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenLink("https://mail.google.com/mail/u/1/?tab=wm&pli=1#imp");
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            OpenLink("https://mail.google.com/mail/u/1/?tab=wm&pli=1#sent");
        }

        #region NavigateEnum
        //iExplorer.Navigate(sUrl, 0x800); //0x800 means new tab
        //navOpenInNewWindow = 0x1, 
        //navNoHistory = 0x2, 
        //navNoReadFromCache = 0x4, 
        //navNoWriteToCache = 0x8, 
        //navAllowAutosearch = 0x10, 
        //navBrowserBar = 0x20, 
        //navHyperlink = 0x40, 
        //navEnforceRestricted = 0x80, 
        //navNewWindowsManaged = 0x0100, 
        //navUntrustedForDownload = 0x0200, 
        //navTrustedForActiveX = 0x0400, 
        //navOpenInNewTab = 0x0800, 
        //navOpenInBackgroundTab = 0x1000, 
        //navKeepWordWheelText = 0x2000, 
        //navVirtualTab = 0x4000, 
        //navBlockRedirectsXDomain = 0x8000, 
        //navOpenNewForegroundTab = 0x10000  
        #endregion


        public void OpenLink(String url)
        {
            try
            {
                // Abre o Browser padrão.
                //System.Diagnostics.Process.Start(url); 

                // Abre o IE.
                //System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo(@"iexplore.exe -framemerging", url);
                //System.Diagnostics.Process.Start(startInfo);
                //startInfo = null;

                // Necessário importar a dll "shdocvw" em Windows\System32.
                ShellWindows iExplorerInstances = new ShellWindows();
                if (iExplorerInstances.Count > 0)
                {
                    IEnumerator enumerator = iExplorerInstances.GetEnumerator();
                    
                    String filename;
                    SHDocVw.InternetExplorer browser;
                    string myLocalLink;
                    mshtml.IHTMLDocument2 myDoc;

                    foreach (InternetExplorer iExplorer in iExplorerInstances)
                    {
                        //if (iExplorer.Name.Contains("Internet Explorer"))
                        filename = System.IO.Path.GetFileNameWithoutExtension(iExplorer.FullName).ToLower();
                        if ((filename == "iexplore"))
                        {
                            enumerator.MoveNext();
                            InternetExplorer iExplorer2 = iExplorer;
                            iExplorer2 = (InternetExplorer)enumerator.Current;
                            {
                                browser = iExplorer;
                                myDoc = (mshtml.IHTMLDocument2)browser.Document;
                                myLocalLink = myDoc.url;

                                if (myLocalLink.Contains("globo"))
                                {
                                    iExplorer2.Navigate(url);
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Não há um Browser IE rodando...
                    MessageBox.Show("Não existe um Internet Explorer aberto.");
                }
                
            }
            catch (Exception ex)
            {
            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            bool found = false;
            ShellWindows iExplorerInstances = new ShellWindows();
            if (iExplorerInstances.Count > 0)
            {
                IEnumerator enumerator = iExplorerInstances.GetEnumerator();
                
                foreach (InternetExplorer iExplorer in iExplorerInstances)
                {
                    if (iExplorer.Name.Contains("Internet Explorer"))
                    {
                        //iExplorer.Navigate(ur, 0x800);
                        found = true;
                        //break;
                    }
                }
                if (!found)
                {
                    //run with processinfo
                }
            }
            MessageBox.Show(found.ToString());
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            SHDocVw.InternetExplorer browser;
            string myLocalLink;
            mshtml.IHTMLDocument2 myDoc;
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();
            string filename;
            foreach (SHDocVw.InternetExplorer ie in shellWindows)
            {
                filename = System.IO.Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
                if ((filename == "iexplore"))
                {
                    browser = ie;
                    myDoc = (mshtml.IHTMLDocument2)browser.Document;
                    myLocalLink = myDoc.url;
                    MessageBox.Show(myLocalLink);
                }
            }
        }

        private void Window_Loaded_1(object sender, RoutedEventArgs e)
        {

        }
    }
}


