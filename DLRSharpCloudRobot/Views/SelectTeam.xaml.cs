using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
//using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DLRSharpCloudRobot.Models;
using DLRSharpCloudRobot.ViewModels;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using CheckBox = System.Windows.Controls.CheckBox;

namespace DLRSharpCloudRobot.Views
{
    /// <summary>
    /// Interaction logic for SelectStory.xaml
    /// </summary>
    public partial class SelectTeam : Window
    {

        public TeamLite SelectedTeam { get; set; }    

        private CollectionViewSource teamsCvs;

        private SelectedStoryViewModel _viewModel;

        private string _userName;

        private bool _allowMultiSelect;
        public SelectTeam(SharpCloudApi client)
        {
            InitializeComponent();

            _viewModel = DataContext as SelectedStoryViewModel;
            _client = client;

            teamsCvs = this.FindResource("teamsCvs") as CollectionViewSource;
            teamsCvs.Source = _viewModel.AllTeams;

            Loaded += new RoutedEventHandler(SelectRoadmap_Loaded);

            SelectedTeam = null;
        }

        public List<string> SelectedIDs { get; private set; }
        public List<StoryLite> SelectedStoryLites { get; private set; }

        private SharpCloudApi _client;


        void SelectRoadmap_Loaded(object sender, RoutedEventArgs e)
        {
            LoadTeams();
        }

        private void OnClickOK(object sender, RoutedEventArgs e)
        {
            if (SelectedTeam == null)
                return;

            DialogResult = true;
        }

        private void OnClickCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void LoadTeams()
        {
            _viewModel.AllTeams.Clear();
            var teams = _client.Teams;
            this.teamStoriesContainer.Visibility = Visibility.Visible;
            foreach (Team t in teams)
            {
                _viewModel.AllTeams.Add(t);
            }

        }

        private void Team_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Border grid = sender as Border;
            if (grid == null) return;
            var item = grid.DataContext as SC.API.ComInterop.Models.Team;
            try
            {
                SelectedTeam = new TeamLite(item);
            }
            catch (Exception e1)
            { }
        }

        private CheckBox _checkBox;


        private void Team_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var a = sender as Border;
            a.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 41, 41, 41));
        }

        private void Team_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var a = sender as Border;
            a.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 204, 22, 76));
        }

        private void imgIcon_Loaded(object sender, RoutedEventArgs e)
        {
            var t = sender as System.Windows.Controls.Image;
            var content = t.DataContext as SC.API.ComInterop.Models.Team;
            String stringPath = _client.BaseUri + "/image/" + content.Id + "?t=teamid";
            Uri imageUri = new Uri(stringPath, UriKind.Absolute);
            BitmapImage imageBitmap = new BitmapImage(imageUri);
            t.Source = imageBitmap;
        }

        private void teamsCvs_Filter(object sender, System.Windows.Data.FilterEventArgs e)
        {
            var userTeam = (SC.API.ComInterop.Models.Team)e.Item;
            if (MatchOnSearchTeam(userTeam.Name))
            {
                e.Accepted = true;
            }
            else
            {
                e.Accepted = false;
            }
        }

        private string searchStrTeam = "";
        private void tbSearchTeam_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.tbSearchTeam.Text != "Search teams")
            {
                searchStrTeam = this.tbSearchTeam.Text.ToLower();
                this.teamsCvs.View.Refresh();
            }
        }

        private bool MatchOnSearchTeam(string str)
        {
            if (!string.IsNullOrEmpty(searchStrTeam))
            {
                str = str.ToLower();
                return str.Contains(searchStrTeam);
            }
            return true;
        }

        private void tbSearchTeam_GotFocus(object sender, RoutedEventArgs e)
        {
            this.tbSearchTeam.Text = "";
            this.tbSearchTeam.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 0, 0));
        }

        private void tbSearchTeam_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.tbSearchTeam.Text == "")
            {
                this.tbSearchTeam.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(194, 194, 194));
                this.tbSearchTeam.Text = "Search teams";
            }
        }

    }
}
