using System.Windows;

namespace OfficeKeys
{
    public partial class AddProductWindow : Window
    {
        public AddProductWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string header = "Accept-Language: en-US\r\n";
            Browser.Navigate("http://setup.office.com/", null, null, header);
        }

        private void Browser_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            string html = null;
            try
            {
                dynamic doc = Browser.Document;
                html = doc.documentElement.InnerHtml;
            }
            catch { }

            if (!string.IsNullOrEmpty(html) && html.Contains("<div id=\"welcomemessageid\">To get started with Office choose Install."))
            {
                DialogResult = true;
            }
        }
    }
}
