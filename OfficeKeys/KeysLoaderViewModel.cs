using Mach.Wpf.Mvvm;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace OfficeKeys
{
    public class KeysLoaderViewModel : NotifyPropertyBase
    {
        private string _userName;
        private string _microsoftAccount;
        private string _cookies;
        private bool _isLoadingInProgress;
        private ObservableCollection<Product> _products;
        private DelegateCommand _loginCommand;
        private DelegateCommand _addProductCommand;
        private DelegateCommand _refreshKeysCommand;
        private DelegateCommand<Product> _copyKeyCommand;

        public string MicrosoftAccount
        {
            get { return _microsoftAccount; }
            set
            {
                _microsoftAccount = value;
                OnPropertyChanged();
            }
        }
        public string UserName
        {
            get { return _userName; }
            set
            {
                _userName = value;
                OnPropertyChanged();
            }
        }

        public string Cookies
        {
            get { return _cookies; }
            set
            {
                Properties.Settings.Default.Cookies = value;
                _cookies = value;

                OnPropertyChanged();
            }
        }

        public bool IsLoadingInProgress
        {
            get { return _isLoadingInProgress; }
            set
            {
                _isLoadingInProgress = value;

                OnPropertyChanged();
            }
        }

        public ObservableCollection<Product> Products
        {
            get { return _products; }
            set
            {
                _products = value;

                OnPropertyChanged();
            }
        }

        public ICommand AddProductCommand
        {
            get { return _addProductCommand; }
        }

        public ICommand RefreshKeysCommand
        {
            get { return _refreshKeysCommand; }
        }

        public ICommand CopyKeyCommand
        {
            get { return _copyKeyCommand; }
        }

        public ICommand LoginCommand
        {
            get { return _loginCommand; }
        }

        public KeysLoaderViewModel()
        {
            _loginCommand = new DelegateCommand(Login);
            _addProductCommand = new DelegateCommand(AddProduct);
            _refreshKeysCommand = new DelegateCommand(RefreshKeys);
            _copyKeyCommand = new DelegateCommand<Product>(CopyKey);
            _products = new ObservableCollection<Product>();
            _lastProducts = new List<Product>();

            Cookies = Properties.Settings.Default.Cookies;

            ChangeWebBrowserVersion();
        }

        private void ChangeWebBrowserVersion()
        {
            string installkey = @"SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
            string entryLabel = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            OperatingSystem osInfo = Environment.OSVersion;

            string version = osInfo.Version.Major.ToString() + '.' + osInfo.Version.Minor.ToString();
            uint editFlag = (uint)0x2710;

            RegistryKey existingSubKey = Registry.CurrentUser.OpenSubKey(installkey, false); // readonly key

            if (existingSubKey.GetValue(entryLabel) == null)
            {
                existingSubKey = Registry.CurrentUser.OpenSubKey(installkey, true); // writable key
                existingSubKey.SetValue(entryLabel, unchecked((int)editFlag), RegistryValueKind.DWord);
            }
        }

        private void AddProduct()
        {
            if (_products == null || _products.Count == 0)
            {
                RefreshKeys();
            }

            AddProductWindow addProduct = new AddProductWindow();
            bool? result = addProduct.ShowDialog();

            if (result ?? true)
            {
                RefreshKeys();
            }
        }

        private void Login()
        {
            LoginWindow login = new LoginWindow();
            bool? result = login.ShowDialog();

            if (result ?? true)
            {
                Cookies = login.AuthorizationCookie;
                MicrosoftAccount = login.MicrosoftAccount;
                UserName = login.UserName;

                RefreshKeys();
            }
        }

        private void CopyKey(Product product)
        {
            System.Windows.Clipboard.SetText(product.Key);
        }

        private List<Product> _loadedProducts;

        private void RefreshKeys()
        {
            IsLoadingInProgress = true;
            _products.Clear();

            Task.Factory.StartNew(() => DoRefreshKeys()).ContinueWith((t) =>
            {
                if (_loadedProducts.Count > 0)
                {
                    _products.Clear();

                    foreach (var product in _loadedProducts)
                    {
                        var previousKey = _lastProducts.Where(p => p.Key == product.Key).FirstOrDefault();
                        if (previousKey == null)
                            product.IsNew = true;

                        _products.Add(product);
                    }

                    _lastProducts = _loadedProducts;
                }

                IsLoadingInProgress = false;

            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private List<Product> _lastProducts;
        private void DoRefreshKeys()
        {
            KeysLoader loader = new KeysLoader();
            _loadedProducts = loader.Load(_cookies);
        }
    }
}
