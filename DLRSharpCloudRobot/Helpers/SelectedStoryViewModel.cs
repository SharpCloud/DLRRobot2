using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using SC.API.ComInterop.Models;

namespace DLRSharpCloudRobot.ViewModels
{
    public class SelectedStoryViewModel : INotifyPropertyChanged
    {
        public SelectedStoryViewModel()
        {
        }

        public string Status
        {
            get { return _status; }
            set
            {
                _status = value;
                OnPropertyChanged("Status");
            }
        }
        private string _status;


        public ICollectionView Stories
        {
            get { return _stories; }
            set
            {
                _stories = value;
                OnPropertyChanged("Stories");
            }
        }
        private ICollectionView _stories;

        public event PropertyChangedEventHandler PropertyChanged;
        private Dispatcher currentDispatcher;

        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                //check if we are on the UI thread if not switch
                if (Dispatcher.CurrentDispatcher.CheckAccess())
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                else
                    Dispatcher.CurrentDispatcher.BeginInvoke(new Action<string>(this.OnPropertyChanged), propertyName);
            }
        }
        public ObservableCollection<StoryLite> AllStories
        {
            get
            {
                if (this.allStories == null)
                    this.allStories = new ObservableCollection<SC.API.ComInterop.Models.StoryLite>();

                return this.allStories;
            }
            set
            {
                if (this.allStories != value)
                {
                    this.allStories = value;
                    OnPropertyChanged("AllStories");
                }
            }
        }
        private ObservableCollection<StoryLite> allStories;

        public ObservableCollection<Team> AllTeams
        {
            get
            {
                if (this.allTeams == null)
                    this.allTeams = new ObservableCollection<SC.API.ComInterop.Models.Team>();

                return this.allTeams;
            }
            set
            {
                if (this.allTeams != value)
                {
                    this.allTeams = value;
                    OnPropertyChanged("AllTeams");
                }
            }
        }
        private ObservableCollection<Team> allTeams;


    }
}
