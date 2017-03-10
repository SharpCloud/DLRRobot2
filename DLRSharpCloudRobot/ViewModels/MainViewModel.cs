using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using DLRSharpCloudRobot.Models;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using SC.Entities.Models;
using Attribute = SC.API.ComInterop.Models.Attribute;
using ModelHelper = DLRSharpCloudRobot.Helpers.ModelHelper;
using Relationship = SC.API.ComInterop.Models.Relationship;

namespace DLRSharpCloudRobot.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public MainViewModel()
        {
            ProgressRange = 100;
        }

        public bool _cancel;

        public string Url { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Proxy { get; set; }
        public bool ProxyAnnonymous { get; set; }
        public string ProxyUserName { get; set; }
        public string ProxyPassword { get; set; }

        public int ProgressRange { get; set; }
        public int ProgressValue { get; set; }

        public string CurrentPeriod
        {
            get { return $"Current Period is {ProcessData.GetCurrentPeriod()}"; }
        }

        public bool RememberPassword
        {
            get { return _rememberPassword; }
            set
            {
                _rememberPassword = value;
                OnPropertyChanged("RememberPassword");
            }
        }
        private bool _rememberPassword;

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


        public List<StoryLite2> TeamStories
        {
            get
            {
                try
                {
                    if (_teamStories == null)
                    {
                        var ts = GetApi().StoriesTeam(SelectedTeam.Id);
                        _teamStories = new List<StoryLite2>();
                        foreach (var t in ts)
                        {
                            if (t.Id != SelectedPortfolioStory.Id && t.Id != SelectedTemplateStory.Id)
                            {
                                var sl2 = new StoryLite2(t);
                                _teamStories.Add(sl2);
                            }
                        }
                        OnPropertyChanged("TeamStories");
                    }
                    return _teamStories;
                }
                catch (Exception e)
                {
                    // swallow
                    return null;
                }
            }
        }

        private List<StoryLite2> _teamStories;

        public void ClearTeamStories()
        {
            _teamStories = null;
            OnPropertyChanged("TeamStories");
        }

        public void RefreshTeamStories()
        {
            OnPropertyChanged("TeamStories");
        }

        public void ClearLogs()
        {
            _logs = string.Empty;
            OnPropertyChanged("Logs");
        }

        public string Logs
        {
            get { return _logs; }
            set
            {
                _logs = value;
                OnPropertyChanged("Logs");
            }
        }

        private string _logs = string.Empty;

        public void AddLog(string str)
        {
            Logs += str + "\r\n";
        }

        public bool ShowWaitForm
        {
            get { return _showWaitForm; }
            set
            {
                _showWaitForm = value;
                OnPropertyChanged("ShowWaitForm");
            }
        }
        private bool _showWaitForm;

        public async void ShowWaitFormNow(string message)
        {
            ShowWaitForm = true;
            Status = message;
            await Task.Delay(10);
        }

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

        public void SaveUserData()
        {
            ModelHelper.RegWrite("Url", Url);
            ModelHelper.RegWrite("UserName", UserName);
            ModelHelper.RegWrite("RememberPassword", RememberPassword.ToString());
            if (RememberPassword)
                ModelHelper.RegWrite("Password2", Convert.ToBase64String(Encoding.Default.GetBytes(Password)));
            ModelHelper.RegWrite("Proxy", Proxy);
            ModelHelper.RegWrite("ProxyAnnonymous", ProxyAnnonymous.ToString());
            ModelHelper.RegWrite("ProxyUserName", ProxyUserName);
            ModelHelper.RegWrite("ProxyPassword", Convert.ToBase64String(Encoding.Default.GetBytes(ProxyPassword)));
            ModelHelper.RegWrite("DataFolder", SelectedDataFolder);
            ModelHelper.RegWrite("Team", ModelHelper.SerializeJSON(SelectedTeam));
            ModelHelper.RegWrite("TemplateStory", ModelHelper.SerializeJSON(SelectedTemplateStory));
            ModelHelper.RegWrite("PortfolioStory", ModelHelper.SerializeJSON(SelectedPortfolioStory));

        }

        public void LoadData()
        {
            ProxyPassword = Encoding.Default.GetString(Convert.FromBase64String(ModelHelper.RegRead("ProxyPassword", "")));
            ProxyUserName = ModelHelper.RegRead("ProxyUserName", "");
            ProxyAnnonymous = Boolean.Parse(ModelHelper.RegRead("ProxyAnnonymous", true.ToString()));
            Proxy = ModelHelper.RegRead("Proxy", "");
            Url = ModelHelper.RegRead("Url", "https://my.sharpcloud.com");
            UserName = ModelHelper.RegRead("UserName", ConfigurationManager.AppSettings["userid"]);
            RememberPassword = ModelHelper.RegRead("RememberPassword", true.ToString()) == true.ToString();
            SelectedDataFolder = ModelHelper.RegRead("DataFolder", ConfigurationManager.AppSettings["WorkingFolder"]);

            if (RememberPassword)
            {
                Password = Encoding.Default.GetString(Convert.FromBase64String(ModelHelper.RegRead("Password2", "")));
            }

            SelectedTeam = ModelHelper.DeserializeJSON<TeamLite>(ModelHelper.RegRead("Team", ModelHelper.SerializeJSON(SelectedTeam)));
            SelectedTemplateStory = ModelHelper.DeserializeJSON<StoryLite2>(ModelHelper.RegRead("TemplateStory", ModelHelper.SerializeJSON(SelectedTemplateStory)));
            SelectedPortfolioStory = ModelHelper.DeserializeJSON<StoryLite2>(ModelHelper.RegRead("PortfolioStory", ModelHelper.SerializeJSON(SelectedPortfolioStory)));

            SaveAllData();
        }

        public string SelectedTeamName
        {
            get { return SelectedTeam.Name; }
        }

        public TeamLite SelectedTeam
        {
            get { return _selectedTeam; }

            set
            {
                _selectedTeam = value;
                OnPropertyChanged("SelectedStory");
                OnPropertyChanged("SelectedTeamName");
            }
        }
        private TeamLite _selectedTeam;

        public string SelectedDataFolder
        {
            get { return _selectedDataFolder; }

            set
            {
                _selectedDataFolder = value;
                OnPropertyChanged("SelectedDataFolder");
            }
        }
        private string _selectedDataFolder;

        public StoryLite2 SelectedTemplateStory
        {
            get { return _selectedTemplateStory; }

            set
            {
                _selectedTemplateStory = value;
                OnPropertyChanged("SelectedTemplateStory");
                OnPropertyChanged("SelectedTemplateName");
            }
        }
        private StoryLite2 _selectedTemplateStory;

        public StoryLite2 SelectedPortfolioStory
        {
            get { return _selectedPortfolioStory; }

            set
            {
                _selectedPortfolioStory = value;
                OnPropertyChanged("SelectedPortfolioStory");
                OnPropertyChanged("SelectedPortfolioName");
            }
        }
        private StoryLite2 _selectedPortfolioStory;

        public string SelectedTemplateName
        {
            get { return SelectedTemplateStory.Name; }
        }
        public string SelectedPortfolioName
        {
            get { return SelectedPortfolioStory.Name; }
        }

        public bool SetupComplete()
        {
            if (string.IsNullOrEmpty(SelectedDataFolder))
                return false;
            if (System.IO.Directory.Exists(SelectedDataFolder))
                return false;


            return true;
        }


        public void SaveAllData()
        {
            SaveUserData();
        }

        public SharpCloudApi GetApi()
        {
            return new SharpCloudApi(UserName, Password, Url, Proxy);
        }

        private bool _bCosts;
        private bool _bRisks;
        private bool _bMilestones;

        public void Process(bool bCosts, bool bRisks, bool bMilestones)
        {
            _bCosts = bCosts;
            _bRisks = bRisks;
            _bMilestones = bMilestones;

            SaveAllData();
            //open up all the stories
            _cancel = false;
            var worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
        }


        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var dlr = new ProcessData(this);
                dlr.ProcessDataNow(true, _bCosts, _bRisks, _bMilestones,
                    TeamStories.Where(ts => ts.IsSelected == true).ToList());
            }
            catch (Exception exception)
            {
                Status = exception.Message;
                AddLog(exception.Message);
            }
            ShowWaitForm = false;
        }

    }
}
