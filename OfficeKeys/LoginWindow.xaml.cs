using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace OfficeKeys
{

    public static class cookies
    {
        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookieEx(string url, string cookieName, StringBuilder cookieData, ref int size, Int32 dwFlags, IntPtr lpReserved);
        private const Int32 InternetCookieHttponly = 0x2000;
        /// <summary> 
        /// Gets the URI cookie container. 
        /// </summary> 
        /// <param name="uri">The URI.</param> 
        /// <returns></returns> 
        public static CookieContainer GetUriCookieContainer(Uri uri)
        {
            CookieContainer cookies = null;
            // Determine the size of the cookie     
            int datasize = 8192 * 16;
            StringBuilder cookieData = new StringBuilder(datasize);
            if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref datasize, InternetCookieHttponly, IntPtr.Zero))
            {
                if (datasize < 0) return null;
                // Allocate stringbuilder large enough to hold the cookie         
                cookieData = new StringBuilder(datasize);

                if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref datasize, InternetCookieHttponly, IntPtr.Zero))
                    return null;
            }

            if (cookieData.Length > 0)
            {
                cookies = new CookieContainer(); cookies.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return cookies;
        }

    }

    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        private string _authorizationCookie;
        private string _userName;
        private string _microsoftAccount;

        public string AuthorizationCookie
        {
            get { return _authorizationCookie; }
        }
        public string UserName
        {
            get { return _userName; }
        }
        public string MicrosoftAccount
        {
            get { return _microsoftAccount; }
        }

        public LoginWindow()
        {
            InitializeComponent();
        }

        private void Browser_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            Uri uri = new Uri("https://stores.office.com/myaccount/home.aspx");
            string coo = Application.GetCookie(uri);

            string html = null;
            try
            {
                dynamic doc = Browser.Document;
                html = doc.documentElement.InnerHtml;
            }
            catch { }

            if (!string.IsNullOrEmpty(html) && html.Contains(">Sign out</a></span></span>"))
            {
                var co = cookies.GetUriCookieContainer(uri);
                _authorizationCookie = co.GetCookieHeader(uri);

                Match m;
                
                m = Regex.Match(html, "<span class=\\\"headMeName\\\">([^<]+)</span>");
                if(m.Success)
                {
                    _userName = m.Groups[1].Value;
                }

                m = Regex.Match(html, "h1 class=\\\"oxmuxh\\\" id=\\\"myaccountHeader\\\">My Office Account&nbsp;&nbsp;-&nbsp;&nbsp;([^<]+)</h1>");
                if (m.Success)
                {
                    _microsoftAccount = m.Groups[1].Value;
                }

                DialogResult = true;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Browser.Navigate("http://office.com/myaccount?ui=en-US");
        }
    }
}
