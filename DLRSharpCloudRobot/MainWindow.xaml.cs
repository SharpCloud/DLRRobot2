using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using DLRSharpCloudRobot.ViewModels;
using DLRSharpCloudRobot.Views;
using SC.API.ComInterop.Models;

namespace DLRSharpCloudRobot
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = DataContext as MainViewModel;
            _viewModel.LoadData();
            tbUrl.Text = _viewModel.Url;
            tbUsername.Text = _viewModel.UserName;
            tbPassword.Password = _viewModel.Password;
            //tbDataFolder.Text = _viewModel.DataFolder;

            if (!ValidateCreds())
                mainTab.SelectedIndex = 0;
            else if (_viewModel.SetupComplete()) // NEED TO ADD SOME 
                mainTab.SelectedIndex = 2;
            else
                mainTab.SelectedIndex = 2;
        }

        

        private void ClickClearPassword(object sender, RoutedEventArgs e)
        {
            tbPassword.Password = "";
            Helpers.ModelHelper.RegWrite("Password2", "");
        }

        private void SaveAndValidateCLick(object sender, RoutedEventArgs e)
        {
            if (ValidateCreds())
            {
                _viewModel.UserName = tbUsername.Text;
                _viewModel.Url = tbUrl.Text;
                _viewModel.Password = tbPassword.Password;

                _viewModel.SaveAllData();
                MessageBox.Show("Well done! Your credentials have been validated.");
            }
            else
            {
                MessageBox.Show("Sorry, your credentials are not correct, please try again.");
            }

        }
        private bool ValidateCreds()
        {
            return SC.API.ComInterop.SharpCloudApi.UsernamePasswordIsValid(tbUsername.Text, tbPassword.Password,
                tbUrl.Text, _viewModel.Proxy, _viewModel.ProxyAnnonymous, _viewModel.ProxyUserName, _viewModel.ProxyPassword);
        }

        private void SelectTeam_Click(object sender, RoutedEventArgs e)
        {
            var sel = new SelectTeam(_viewModel.GetApi());
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedTeam = sel.SelectedTeam;
            }
            _viewModel.SaveAllData();
        }

        private void SelectPortfolio_Click(object sender, RoutedEventArgs e)
        {
            var sel = new SelectStory(_viewModel.GetApi(), false);
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedPortfolioStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
            }
            _viewModel.SaveAllData();
        }

        private void SelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sel = new SelectStory(_viewModel.GetApi(), false);
                if (sel.ShowDialog() == true)
                {
                    _viewModel.SelectedTemplateStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
                }
            }
            catch (Exception E)
            {

            }
            _viewModel.SaveAllData();
        }

        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                _viewModel.SelectedDataFolder = dialog.SelectedPath;
            }
            _viewModel.SaveAllData();
        }

        private void SelectConfigFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", System.AppDomain.CurrentDomain.BaseDirectory);
        }

        private void StartProcess_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.Process((bool)chCosts.IsChecked, (bool)chRisks.IsChecked, (bool)chMilestones.IsChecked);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _viewModel._cancel = true;
        }

        private void Button_ClickSelectAll(object sender, RoutedEventArgs e)
        {
            foreach (var s in _viewModel.TeamStories)
                s.IsSelected = true;
            teamStoriesChkList.ItemsSource = null;
            teamStoriesChkList.ItemsSource = _viewModel.TeamStories;
        }

        private void Button_ClickRefresh(object sender, RoutedEventArgs e)
        {
            _viewModel.ClearTeamStories();
        }

        private void Button_ClickSelectViewLogs(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", _viewModel.SelectedDataFolder);
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var tb = sender as TextBox;
            tb.ScrollToEnd();
        }
    }
}
