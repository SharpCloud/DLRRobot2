using System;
using System.Collections.Generic;
using System.Globalization;
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
using System.Windows.Shapes;
using DLRSharpCloudRobot.ViewModels;

namespace DLRSharpCloudRobot.Views
{
    /// <summary>
    /// Interaction logic for ProxySettings.xaml
    /// </summary>
    public partial class ProxySettings : Window
    {
        private MainViewModel _viewModel;

        public ProxySettings(MainViewModel viewModel)
        {
            _viewModel = viewModel;

            InitializeComponent();

            tbProxy.Text = _viewModel.Proxy;
            chkAnnonymous.IsChecked = _viewModel.ProxyAnnonymous;
            tbUsername.Text = _viewModel.ProxyUserName;
            tbPassword.Password = _viewModel.ProxyPassword;

            tbPassword.Password = _viewModel.Password;

        }

        private void ClickOnOK(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(tbProxy.Text) && (bool)!chkAnnonymous.IsChecked)
            {
                if (string.IsNullOrEmpty(tbUsername.Text) || string.IsNullOrEmpty(tbPassword.Password))
                {
                    MessageBox.Show(
                        "You must provide a username and password if you are not using an anonymous proxy.",
                        "Proxy Server error");
                    return;
                }
            }

            _viewModel.Proxy = tbProxy.Text;
            _viewModel.ProxyAnnonymous = (bool)chkAnnonymous.IsChecked;
            _viewModel.ProxyUserName = tbUsername.Text;
            _viewModel.ProxyPassword = tbPassword.Password;
            _viewModel.SaveAllData();
            Close();
        }

        private void ClickCancel(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
