using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Net;
using System.Security.Policy;
using System.Threading.Tasks;
using DLRSharpCloudRobot.Models;
using DLRSharpCloudRobot.ViewModels;
using Attribute = SC.API.ComInterop.Models.Attribute;
using Directory = System.IO.Directory;

namespace DLRSharpCloudRobot
{
    public class ProcessData
    {
        private static int _period;
        private static int _periods;
        private static bool _firstDayOfPeriod;
        private static string _log;
        private static int _errCount;
        private static DateTime[] _periodDates;
        private static List<MilestoneData> _milesStones;
        private static string _workingFolder;
        private static string _logFile;

        private MainViewModel _vm;

        [DllImport("user32.dll")]
        public static extern int FindWindow(string strclassName, string strWindowName);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void KillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr) hWnd, out processID);
            if (processID == 0)
                throw new ArgumentException("Process has not been found by the given main window handle.", "hWnd");
            Process.GetProcessById((int) processID).Kill();
        }

        private static void CloseAllExcels()
        {
            int hwnd = FindWindow("XLMAIN", null);

            while (hwnd != 0)
            {
                KillProcessByMainWindowHwnd(hwnd);
                hwnd = FindWindow("XLMAIN", null);
            }

        }

        public ProcessData(MainViewModel vm)
        {
            _vm = vm;            
        }


        public void ProcessDataNow(bool bKillExcels, bool bCosts, bool bRisks, bool bMilestones, List<StoryLite2> stories)
        {
            _vm.ShowWaitForm = true;
            try
            {
                _workingFolder = _vm.SelectedDataFolder;// ConfigurationManager.AppSettings["WorkingFolder"];
                // check folder exists
                if (!Directory.Exists(_workingFolder))
                {
                    LogError($"Directory '{_workingFolder}' dos not exist - exiting!");
                }
                var logPath = $"{_workingFolder}/logs";
                if (!Directory.Exists(logPath))
                    Directory.CreateDirectory(logPath);
                var now = DateTime.Now;
                _logFile = $"{logPath}/{now.Year}-{now.Month}-{now.Day}-log.txt";

            }
            catch (Exception e)
            {
                _vm.AddLog("Unable to setup working folder");
                _vm.AddLog(e.Message);
                _vm.ShowWaitForm = false;
                return;
            }

            // kill any remaining Exel sessions 
            //var bKillExcels = ConfigurationManager.AppSettings["PreventKillExcelsAtStart"];
            //if (!string.IsNullOrEmpty(bKillExcels) && bKillExcels.ToLower().Trim() != "true")
            try
            {
                if (bKillExcels)
                    CloseAllExcels();
            }
            catch (Exception eke)
            {
                Log(eke.Message);
            }
            //GetStoryPanels();
            // SetStoryPanels();
            //return;

            _periods = Int32.Parse(ConfigurationManager.AppSettings["numberOfPeriods"]);
            _period = GetCurrentPeriod();
            _firstDayOfPeriod = IsPerdiodFirstDay();
            LoadPeriods(_periods);

            var userid = _vm.UserName;//ConfigurationManager.AppSettings["userid"];
            var passwd = _vm.Password;//ConfigurationManager.AppSettings["passwd"];

            var sc = new SharpCloudApi(userid, passwd);

            var teamid = _vm.SelectedTeam.Id;// ConfigurationManager.AppSettings["teamid"];
            var portfolio = _vm.SelectedPortfolioStory.Id;// ConfigurationManager.AppSettings["portfolioid"];
            var templateId = _vm.SelectedTemplateStory.Id;// ConfigurationManager.AppSettings["templateid"];
            var teststoryid = ConfigurationManager.AppSettings["teststoryid"];

            var startTime = DateTime.UtcNow;
            Log("-------------------------------------------------");
            Log(string.Format("Process started at {0}", startTime.ToLongTimeString()));
            Log(string.Format("Working period is {0}", _period));
            if (_firstDayOfPeriod)
                Log(string.Format("Today is the first day of period '{0}'", _period));


            var portfolioStory = sc.LoadStory(portfolio);
            Log($"Reading from '{portfolioStory.Name}'");

            var XL1 = new Application();
            Workbook wbMilestones = null, wbCosts = null;
            
            var pathMlstn = $"{_workingFolder}/data/Milestones/P{_period:D2}/{ConfigurationManager.AppSettings["MilestoneXLSFilename"]}";
            //var pathMlstn = string.Format(ConfigurationManager.AppSettings["MilestoneXLSLocation"], _period);
            //            var pathCosts = $"{_workingFolder}/data/Costs/{ConfigurationManager.AppSettings["BaselineCostsXLSFilename"]}";
            //var pathCosts = ConfigurationManager.AppSettings["BaselineCostsXLSLocation"];

            if (bMilestones)
            {
                if (File.Exists(pathMlstn))
                {
                    Log($"Opening Excel Doc " + pathMlstn);
                    wbMilestones = XL1.Workbooks.Open(pathMlstn);
                    LoadMilestones(XL1, wbMilestones);
                }
                else
                    Log("Could not find file " + pathMlstn);
            }

            try
            {
                int counter = 1;
                _vm.ProgressRange = stories.Count();
                _vm.ProgressValue = 0;

                foreach (var teamStory in stories)
                {
                    _vm.ProgressValue++;
                    Task.Delay(1000);
                    if (_vm._cancel)
                    {
                        Log($"Cancelling!");
                        _vm.ShowWaitForm = false;
                        return;
                    }

                    if (teamStory.Id != portfolio && teamStory.Id != templateId)
                    {
                        if (string.IsNullOrEmpty(teststoryid) || teamStory.Id == teststoryid)
                        {
                            try
                            {
                                Log($"{counter++} Reading from '{teamStory.Name}'");
                                var sc2 = new SharpCloudApi(userid, passwd);
                                var projectStory = sc2.LoadStory(teamStory.Id);

                                // find the item in the project story 
                                var projectStatusItemName =
                                    string.Format(ConfigurationManager.AppSettings["rollupMainStatusItem"], _period);
                                if (!string.IsNullOrEmpty(projectStatusItemName))
                                {
                                    var projectStatusItem = projectStory.Item_FindByName(projectStatusItemName);
                                    if (projectStatusItem != null)
                                    {
                                        LoadPeriodDates(projectStory);

                                        if (_firstDayOfPeriod)
                                        {
                                            CopyAllItemDataToNextPeriod(projectStory);
                                            CopySpecificItemDataToNextPeriod(projectStory);
                                        }

                                        var portfolioItemID =
                                            projectStatusItem.GetAttributeValueAsText(GetAttribute(projectStory,
                                                "rollupMainAttribute"));
                                        if (bCosts)
                                            ImportCosts(projectStory, portfolioItemID);
                                        //ImportBaseslineCosts(XL2, wbCosts, projectStory, portfolioItemID);
                                        if (bMilestones)
                                            ImportMilestones(projectStory, portfolioItemID);
                                        if (bRisks)
                                            ImportRisks(projectStory, portfolioItemID);

                                        RunPendingStoryChanges(projectStory);
                                        RunStoryCalcs(projectStory);
                                        RunStoryCalcsForNazir(projectStory);
                                        RunEFCCals(projectStory);

                                        SetStoryPanels(projectStory);

                                        SaveStory(projectStory);

                                        if (!string.IsNullOrEmpty(portfolioItemID))
                                        {
                                            var portfolioItem = portfolioStory.Item_FindByExternalId(portfolioItemID);

                                            if (portfolioItem != null)
                                            {
                                                int count =
                                                    int.Parse(ConfigurationManager.AppSettings["rollupItemCount"]);
                                                for (int i = 1; i <= count; i++)
                                                {
                                                    // copy across value
                                                    CopyAttributeValues(portfolioStory, portfolioItem, projectStory,
                                                        ConfigurationManager.AppSettings[
                                                            string.Format("rollupItem{0}", i)]);
                                                }
                                                // make sure the portfolio Item is linked to the right story
                                                var res =
                                                    portfolioItem.Resource_FindByName(
                                                        ConfigurationManager.AppSettings["LinkedProjectName"]);
                                                if (res == null)
                                                {
                                                    res =
                                                        portfolioItem.Resource_AddName(
                                                            ConfigurationManager.AppSettings["LinkedProjectName"]);
                                                }
                                                res.Description = ConfigurationManager.AppSettings["LinkedProjectDesc"];
                                                res.Url = new Uri(projectStory.Url);
                                            }
                                            else
                                            {
                                                LogError(string.Format(
                                                    "71 Could not find item '{0}' by ExternalID in portfolio story '{1}'",
                                                    portfolioItemID, portfolioStory.Name));
                                            }
                                        }
                                        else
                                        {
                                            LogError(string.Format(
                                                "72 Item '{0}' in '{1}' does not have a value a for '{2}'",
                                                projectStatusItem,
                                                projectStory.Name,
                                                ConfigurationManager.AppSettings["rollupMainAttribute"]));
                                        }
                                    }
                                    else
                                    {
                                        LogError(string.Format(
                                            "73 Could not find item called '{0}' in '{1}'", projectStatusItemName,
                                            projectStory.Name));

                                    }
                                }
                                else
                                {
                                    LogError(string.Format("74 No value set in the config for 'projectStatusItem'"));
                                }
                            }
                            catch (Exception e)
                            {
                                LogError(string.Format($"999 failed to load story {teamStory.Name} {e.Message}"));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
            }
            SaveStory(portfolioStory);

            // close Excel down
            if (wbMilestones != null)
                wbMilestones.Close(false);
            if (wbCosts != null)
                wbCosts.Close(false);
            XL1.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(XL1);
            XL1 = null;
            //XL2.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(XL2);
            //XL2 = null;

            // complete teh log file
            var endTime = DateTime.UtcNow;
            Log(string.Format("Process completed at {0}", endTime.ToLongTimeString()));
            Log(string.Format("Process took {0} seconds", (endTime - startTime).TotalSeconds));

            Log(string.Format("There were {0} errors", _errCount));
            // any errors to report?
            if (_errCount > 0)
            {
                SendEmail();
            }
            //Console.ReadKey();
        }

        private void LoadPeriods(int periods)
        {
            _periodDates = new DateTime[_periods + 2];
            _periodDates[0] = DateTime.MinValue;

            try
            {
                for (int i = 1; i <= periods + 1; i++)
                {
                    _periodDates[i] = DateTime.Parse(ConfigurationManager.AppSettings[string.Format("Period{0:D2}", i)]);
                }
            }
            catch (Exception exception)
            {
                throw new Exception("Error initialising current periods");
            }
        }

        private void SendEmail()
        {
            try
            {
                var host = ConfigurationManager.AppSettings["SmtpHost"].Split(':');
                using (SmtpClient client = new SmtpClient(host[0]))
                {
                    client.Port = Int32.Parse(host[1]);

                    try
                    {
                        string test = ConfigurationManager.AppSettings["UsePasswordUsername"];
                        if (test.ToLower() == "true")
                            client.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["SmtpUsername"], ConfigurationManager.AppSettings["SmtpPass"]);
                        else
                            client.Credentials = new NetworkCredential("ac46338", "yaxp7fuxu");
                    }
                    catch (Exception ex)
                    {
                        client.Credentials = new NetworkCredential("ac46338", "yaxp7fuxu");
                    }
                    MailAddress from;
                    from = new MailAddress(ConfigurationManager.AppSettings["FromEmail"], ConfigurationManager.AppSettings["FromName"], System.Text.Encoding.UTF8);
                    MailAddress to = new MailAddress(ConfigurationManager.AppSettings["ErrorEmails"]);
                    using (MailMessage message = new MailMessage(from, to))
                    {
                        message.Subject = "DLR robot had an error";
                        message.Body = _log;
                        message.IsBodyHtml = false;

                        client.Send(message);
                    }
                }
            }
            catch (Exception ex)
            {
                //email has failed so log using trace as LogException error uses email and so would cause probs. This can be changed once we stop emailing exceptions
                LogError("75 Send Email Failed " + ex.Message);
            }
        }


        private void LogError(string str)
        {
            _errCount++;
            Log(@"Error: " + str);
        }

        private void Log(string str)
        {
            str = string.Format("{0} {1}   {2} ", DateTime.UtcNow.ToShortDateString(), DateTime.UtcNow.ToLongTimeString(), str);

            _log += str + "\n";
            Console.WriteLine(str);

            // File.AppendAllText(ConfigurationManager.AppSettings["LogFile"], str + "\r\n");
            File.AppendAllText(_logFile, str + "\r\n");
            _vm.AddLog(str);
        }

        private void SaveStory(Story story)
        {
            try
            {
                if (story.IsModified)
                {
                    Log($"Saving changes '{story.Name}'");
                    story.Description = string.Format("Modified by API at {0} {1}", DateTime.UtcNow.ToShortDateString(),
                        DateTime.UtcNow.ToShortTimeString());
                    story.Save();
                }
                else
                {
                    Log($"No changes detected to '{story.Name}'");
                }
            }
            catch (Exception e)
            {
                LogError($"99 Problem Saving {story.Name} - {e.ToString()}");
            }

        }

        private void CopyAttributeValues(Story portfolio, Item portolioItem, Story projectStory, string text)
        {
            // find the item
            int period = _period - 1; // portfolio looks at prior period
            if (period == 0) // i
                period = 1;
            string itemConfigName = text + "StatusItem";
            string itemName = string.Format(ConfigurationManager.AppSettings[itemConfigName], period);
            var item = projectStory.Item_FindByName(itemName);
            if (item == null)
            {
                LogError(string.Format("76 Could not find item '{0}' in '{1}'", itemName, projectStory.Name));
                return;
            }

            // find the number of items to copy
            int count;
            string attConfigTextCount = text + "AttributesCount";
            if (int.TryParse(ConfigurationManager.AppSettings[attConfigTextCount], out count))
            {
                for (int a = 0; a < count; a++)
                {
                    string attConfigText = string.Format("{0}Attribute{1}", text, a + 1);
                    var attr = GetAttribute(portfolio, attConfigText);
                    if (attr.Type == SC.API.ComInterop.Models.Attribute.AttributeType.Numeric)
                    {
                        portolioItem.SetAttributeValue(attr,
                            item.GetAttributeValueAsDouble(GetAttribute(projectStory, attConfigText)));
                    }
                    else
                    {
                        portolioItem.SetAttributeValue(attr,
                            item.GetAttributeValueAsText(GetAttribute(projectStory, attConfigText)));
                    }
                }
            }
            else
            {
                LogError("77 There was a problem with " + attConfigTextCount);
            }
        }

        private bool IsPerdiodFirstDay()
        {
            try
            {
                var now = DateTime.UtcNow.Date;

                int period = -1; // undefined

                for (int i = 0; i < _periods; i++)
                {
                    var date = DateTime.Parse(ConfigurationManager.AppSettings[string.Format("Period{0:D2}", i + 1)]);
                    if (now == date)
                        return true;
                }
                return false;
            }
            catch (Exception exception)
            {
                throw new Exception("Error initialising dates from Config");
            }
        }

        public static int GetCurrentPeriod()
        {
            try
            {
                var now = DateTime.UtcNow;

                int period = -1; // undefined
                _periods = Int32.Parse(ConfigurationManager.AppSettings["numberOfPeriods"]);

                for (int i = 1; i <= _periods; i++)
                {
                    var date = DateTime.Parse(ConfigurationManager.AppSettings[string.Format("Period{0:D2}", i)]);
                    if (now > date)
                        period = i;
                }
                return period;
            }
            catch (Exception exception)
            {
                throw new Exception("Error initialising current period");
            }
        }

        private void LoadPeriodDates(Story story)
        {
            try
            {
                int itemsToCopy = int.Parse(ConfigurationManager.AppSettings["itemPeriodDateCount"]);
                var attStart = GetAttribute(story, "periodStart");
                var attEnd = GetAttribute(story, "periodEnd");

                for (int i = 1; i <= itemsToCopy; i++)
                {
                    string namePattern = ConfigurationManager.AppSettings[string.Format("itemPeriodDate{0}", i)];

                    for (int p = 1; p <= _periods; p++)
                    {
                        var item = story.Item_FindByName(string.Format(namePattern, p));

                        item.SetAttributeValue(attStart, _periodDates[p]);
                        item.SetAttributeValue(attEnd, _periodDates[p + 1].AddDays(-1));
                    }
                }
            }
            catch (Exception exception)
            {
                LogError(string.Format("LoadPeriodDates: {0}", exception.Message));
            }
        }

        private void CopyAllItemDataToNextPeriod(Story story)
        {
            try
            {
                int itemsToCopy = int.Parse(ConfigurationManager.AppSettings["itemCopyAcrossCount"]);

                for (int i = 1; i <= itemsToCopy; i++)
                {
                    string namePattern = ConfigurationManager.AppSettings[string.Format("itemCopyAcross{0}", i)];

                    var newItem = story.Item_FindByName(string.Format(namePattern, _period));
                    var lastItem = story.Item_FindByName(string.Format(namePattern, _period - 1));

                    if (newItem != null && lastItem != null)
                    {
                        // copy accoss all the attribues
                        foreach (var attribute in story.Attributes)
                        {
                            if (lastItem.GetAttributeIsAssigned(attribute))
                            {
                                newItem.SetAttributeValue(attribute, lastItem.GetAttributeValueAsText(attribute));
                            }
                        }

                    }
                    else
                    {
                        Log(string.Format($"Could not find items for {namePattern}"));
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("78 CopyAllItemDataToNextPeriod: " + ex.Message);
            }
        }

        private void CopySpecificItemDataToNextPeriod(Story story)
        {
            try
            {
                string namePattern = ConfigurationManager.AppSettings["itemCarryOver"];

                var newItem = story.Item_FindByName(string.Format(namePattern, _period));
                var lastItem = story.Item_FindByName(string.Format(namePattern, _period - 1));

                if (newItem != null && lastItem != null)
                {
                    int attribsToCopy = int.Parse(ConfigurationManager.AppSettings["itemCarryOverAttributeCount"]);

                    // copy accoss all the attribues
                    for (int a = 1; a <= attribsToCopy; a++)
                    {
                        var attribute = GetAttribute(story, string.Format("itemCOA{0:D2}", a));
                        if (lastItem.GetAttributeIsAssigned(attribute))
                        {
                            newItem.SetAttributeValue(attribute, lastItem.GetAttributeValueAsText(attribute));
                        }
                    }

                }
                else
                {
                    Log(string.Format($"Could not find items for {namePattern}"));
                }
            }
            catch (Exception ex)
            {
                LogError("79 CopySpecificItemDataToNextPeriod: " + ex.Message);
            }
        }
        
        private void RunPendingStoryChanges(Story story)
        {
            try
            {
                var cat = ConfigurationManager.AppSettings["PendingChangesCategory"];
                var pid = ConfigurationManager.AppSettings["PendingChangesEID"];
                var isPending = ConfigurationManager.AppSettings["PendingChangesAttributePending"];
                var tag = string.Format("P{0:D2}", _period);

                var att = GetAttribute(story, "PendingChangesAttribute");

                foreach (var item in story.Items)
                {
                    if (item.Category.Name == cat) // it is a pending item
                    {
                        if (item.ExternalId != pid)
                        {
                            // its a real pending item
                            if (item.GetAttributeValueAsText(att) == isPending)
                            {
                                item.Tag_AddNew(tag); // make sure our item is tagged for this period - the current period will be taken care of later
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("80 RunPendingStoryChanges: " + ex.Message);
            }
        }

        private void RunStoryCalcsForNazir(Story story)
        {
            var itemFinalYear = story.Item_FindByName(ConfigurationManager.AppSettings["NFYearItem"]);
            if (itemFinalYear == null)
            {
                LogError("93 Final year Financial Item not found");
                return;
            }

            var NFForecast = GetAttribute(story, "NFForecast");
            var NFOverSpend = GetAttribute(story, "NFOverSpend");
            var NFRAG = GetAttribute(story, "NFRAG");
            var NFYearBudget = GetAttribute(story, "NFYearBudget");

            for (int i = 1; i <= 13; i++)
            {
                var itemName = string.Format(ConfigurationManager.AppSettings["NFPeriodItem"], i);

                var item = story.Item_FindByName(itemName);
                if (item != null)
                {
                    var fCast = item.GetAttributeValueAsDouble(NFForecast);
                    var bBudget = itemFinalYear.GetAttributeValueAsDouble(NFYearBudget);

                    var oSpend = fCast - bBudget;

                    if (!item.GetAttributeIsAssigned(NFForecast) || !itemFinalYear.GetAttributeIsAssigned(NFYearBudget))
                        oSpend = 0; // override and set to zero


                    item.SetAttributeValue(NFOverSpend, oSpend);
                    item.SetAttributeValue(NFRAG, GetBudgetVarianceLabelText(GetBudgetVarianceLabelIndex(oSpend, bBudget)));

                }
                else
                {
                    LogError($"93 Item '{itemName}' not found");
                }
            }
        }

        private void RunStoryCalcs(Story story)
        {
            // make sure the current tags are up to date -  we do this every day to make sure things are not broken
            var tagCurrent = story.ItemTag_FindByName(ConfigurationManager.AppSettings["currentPeriod"]);
            if (tagCurrent == null)
            {
                LogError("81 Current Period Tag not found");
            }

            var tagPeriod = story.ItemTag_FindByName(string.Format("P{0:D2}", _period));
            foreach (var item in story.Items)
            {
                if (item.Category.Name != "Dashboard")
                {
                    if (item.Tag_FindById(tagPeriod.Id) != null)
                        item.Tag_AddNew(tagCurrent);
                    else
                        item.Tag_DeleteByTag(tagCurrent);
                }
            }

            // add up cumulative data
            var itemPriorYear = story.Item_FindByName(ConfigurationManager.AppSettings["PriorYearItem"]);
            if (itemPriorYear == null)
            {
                LogError("82 Prior year Financial Item not found");
            }
            var itemFinalYear = story.Item_FindByName(ConfigurationManager.AppSettings["FinalYearItem"]);
            if (itemFinalYear == null)
            {
                LogError("83 Final year Financial Item not found");
            }
            var itemFinalMonth = story.Item_FindByName(ConfigurationManager.AppSettings["LastMonthItem"]);
            if (itemFinalMonth == null)
            {
                LogError("84 Last month Financial Item not found");
            }

            var itemPeriodNames = ConfigurationManager.AppSettings["ItemPeriodNames"];

            var attr = new SC.API.ComInterop.Models.Attribute[19];
            for (int i = 1; i <= 18; i++) // one based
            {
                attr[i] = GetAttribute(story, string.Format("BFC{0:D2}", i));
            }

            // run calcs on priorYear
            itemPriorYear.SetAttributeValue(attr[2], itemPriorYear.GetAttributeValueAsDouble(attr[1]));
            itemPriorYear.SetAttributeValue(attr[3], itemPriorYear.GetAttributeValueAsDouble(attr[1]));
            itemPriorYear.SetAttributeValue(attr[5], itemPriorYear.GetAttributeValueAsDouble(attr[4]));
            itemPriorYear.SetAttributeValue(attr[6], itemPriorYear.GetAttributeValueAsDouble(attr[4]));
            itemPriorYear.SetAttributeValue(attr[7], itemPriorYear.GetAttributeValueAsDouble(attr[4]) - itemPriorYear.GetAttributeValueAsDouble(attr[1]));
            itemPriorYear.SetAttributeValue(attr[8], itemPriorYear.GetAttributeValueAsDouble(attr[5]) - itemPriorYear.GetAttributeValueAsDouble(attr[2]));

            itemPriorYear.SetAttributeValue(attr[17],
                GetCostVarianceLabelText(
                    GetCostVarianceLabelIndex(itemPriorYear.GetAttributeValueAsDouble(attr[7]), itemPriorYear.GetAttributeValueAsDouble(attr[1]))));
            itemPriorYear.SetAttributeValue(attr[18],
                GetCostVarianceLabelText(
                    GetCostVarianceLabelIndex(itemPriorYear.GetAttributeValueAsDouble(attr[8]), itemPriorYear.GetAttributeValueAsDouble(attr[2]))));


            double ytdBaseLineTotal = itemPriorYear.GetAttributeValueAsDouble(attr[1]);
            double lastBaseLineTotal = 0;
            double ytdTotal = itemPriorYear.GetAttributeValueAsDouble(attr[4]);
            double lastTotal = 0;
            for (int p = 1; p <= _periods; p++)
            {
                var itemName = string.Format(itemPeriodNames, p);
                var itemFinancial = story.Item_FindByName(itemName);
                if (itemFinancial != null)
                {
                    // Baseline
                    double baselineValue = itemFinancial.GetAttributeValueAsDouble(attr[1]);
                    lastBaseLineTotal += baselineValue;
                    itemFinancial.SetAttributeValue(attr[2], lastBaseLineTotal);
                    itemFinancial.SetAttributeValue(attr[3], lastBaseLineTotal + ytdBaseLineTotal);

                    // Projection
                    double value = itemFinancial.GetAttributeValueAsDouble(attr[4]);
                    lastTotal += value;
                    itemFinancial.SetAttributeValue(attr[5], lastTotal);
                    itemFinancial.SetAttributeValue(attr[6], lastTotal + ytdTotal);

                    // Variance
                    double mnth = value - baselineValue;
                    double mnthCum = lastTotal - lastBaseLineTotal;
                    // force ranges until SC supports custom buckets
                    if (mnth > 500000) mnth = 500000;
                    if (mnth < -500000) mnth = -500000;
                    if (mnthCum > 500000) mnthCum = 500000;
                    if (mnthCum < -500000) mnthCum = -500000;

                    itemFinancial.SetAttributeValue(attr[7], mnth);
                    itemFinancial.SetAttributeValue(attr[8], mnthCum);

                    itemFinancial.SetAttributeValue(attr[17],
                        GetCostVarianceLabelText(
                            GetCostVarianceLabelIndex(itemFinancial.GetAttributeValueAsDouble(attr[7]), itemFinancial.GetAttributeValueAsDouble(attr[1]))));
                    itemFinancial.SetAttributeValue(attr[18],
                        GetCostVarianceLabelText(
                            GetCostVarianceLabelIndex(itemFinancial.GetAttributeValueAsDouble(attr[8]), itemFinancial.GetAttributeValueAsDouble(attr[2]))));

                }
                else
                {
                    LogError("85 Could not find item:" + itemName);
                }
            }

            // final year item
            itemFinalYear.SetAttributeValue(attr[9], itemFinalMonth.GetAttributeValueAsDouble(attr[2]));
            itemFinalYear.SetAttributeValue(attr[10], itemFinalMonth.GetAttributeValueAsDouble(attr[5]));
            itemFinalYear.SetAttributeValue(attr[13], itemFinalYear.GetAttributeValueAsDouble(attr[12]) - itemFinalYear.GetAttributeValueAsDouble(attr[11]));
            itemFinalYear.SetAttributeValue(attr[14], itemFinalYear.GetAttributeValueAsDouble(attr[9]) + itemFinalYear.GetAttributeValueAsDouble(attr[11]) + itemPriorYear.GetAttributeValueAsDouble(attr[1]));
            itemFinalYear.SetAttributeValue(attr[15], itemFinalYear.GetAttributeValueAsDouble(attr[10]) + itemFinalYear.GetAttributeValueAsDouble(attr[12]) + itemPriorYear.GetAttributeValueAsDouble(attr[4]));
            itemFinalYear.SetAttributeValue(attr[16], itemFinalYear.GetAttributeValueAsDouble(attr[15]) - itemFinalYear.GetAttributeValueAsDouble(attr[14]));



            // new portfolio status item

        }

        private void RunEFCCals(Story story)
        {
            try
            {

                var itemName = string.Format(ConfigurationManager.AppSettings["EFCItemName"], _period);
                var itemNameEFCPrior = string.Format(ConfigurationManager.AppSettings["ItemPeriodNames"], _period - 1);

                var item = story.Item_FindByName(itemName);
                var itemPrior = story.Item_FindByName(itemNameEFCPrior);

                if (itemPrior == null)
                {
                    LogError(string.Format("86 Could not find '{0}'", itemPrior));
                }


                if (item != null)
                {
                    var attrAuth = new SC.API.ComInterop.Models.Attribute[8];
                    var attrValsAuth = new double[8];
                    for (int i = 1; i <= 7; i++) // one based
                    {
                        attrAuth[i] = GetAttribute(story, string.Format("AUTH{0:D2}", i));
                        attrValsAuth[i] = item.GetAttributeValueAsDouble(attrAuth[i]);
                    }

                    // set up our attributes
                    var attr = new SC.API.ComInterop.Models.Attribute[33];
                    var attrVals = new double[33];
                    for (int i = 1; i <= 32; i++) // one based
                    {
                        attr[i] = GetAttribute(story, string.Format("EFC{0:D2}", i));
                        attrVals[i] = item.GetAttributeValueAsDouble(attr[i]);
                    }

                    //calcs
                    attrVals[3] = attrVals[1] - attrVals[2];
                    attrVals[6] = attrVals[4] - attrVals[5];
                    attrVals[7] = attrVals[1] + attrVals[4];
                    attrVals[8] = attrVals[2] + attrVals[5];
                    attrVals[9] = attrVals[3] + attrVals[6];
                    attrVals[12] = attrVals[10] - attrVals[11];
                    attrVals[15] = attrVals[13] - attrVals[14];
                    attrVals[18] = attrVals[16] - attrVals[17];
                    attrVals[19] = attrVals[10] + attrVals[13] + attrVals[16];
                    attrVals[20] = attrVals[11] + attrVals[14] + attrVals[17];
                    attrVals[21] = attrVals[12] + attrVals[15] + attrVals[18];
                    attrVals[22] = attrVals[7] - attrVals[19];
                    attrVals[23] = attrVals[8] - attrVals[20];
                    attrVals[24] = attrVals[22] - attrVals[23];
                    attrVals[25] = item.GetAttributeValueAsDouble(GetAttribute(story, "BFC06"));
                    if (itemPrior != null)
                    {
                        attrVals[26] = itemPrior.GetAttributeValueAsDouble(GetAttribute(story, "BFC06"));
                    }
                    attrVals[27] = attrVals[25] - attrVals[26];
                    if (attrVals[7] != 0) // check for divide be zero
                        attrVals[28] = (attrVals[25] / attrVals[7]) * 100;
                    if (attrVals[8] != 0)
                        attrVals[29] = (attrVals[26] / attrVals[8]) * 100;
                    attrVals[30] = attrVals[28] - attrVals[29];


                    // set values
                    item.SetAttributeValue(attr[3], attrVals[3]);
                    item.SetAttributeValue(attr[6], attrVals[6]);
                    item.SetAttributeValue(attr[7], attrVals[7]);
                    item.SetAttributeValue(attr[8], attrVals[8]);
                    item.SetAttributeValue(attr[9], attrVals[9]);
                    item.SetAttributeValue(attr[12], attrVals[12]);
                    item.SetAttributeValue(attr[15], attrVals[15]);
                    item.SetAttributeValue(attr[18], attrVals[18]);
                    item.SetAttributeValue(attr[19], attrVals[19]);
                    item.SetAttributeValue(attr[20], attrVals[20]);
                    item.SetAttributeValue(attr[21], attrVals[21]);
                    item.SetAttributeValue(attr[22], attrVals[22]);
                    item.SetAttributeValue(attr[23], attrVals[23]);
                    item.SetAttributeValue(attr[24], attrVals[24]);
                    item.SetAttributeValue(attr[25], attrVals[25]);
                    item.SetAttributeValue(attr[26], attrVals[26]);
                    item.SetAttributeValue(attr[27], attrVals[27]);
                    item.SetAttributeValue(attr[28], attrVals[28]);
                    item.SetAttributeValue(attr[29], attrVals[29]);
                    item.SetAttributeValue(attr[30], attrVals[30]);

                    item.SetAttributeValue(attr[31], GetFinanceVarianceLabelText(GetFinaceVarianceLabelIndex(attrVals[22])));
                    item.SetAttributeValue(attr[32], GetFinanceVarianceLabelText(GetFinaceVarianceLabelIndex(attrVals[23])));


                    attrValsAuth[4] = attrVals[7] - attrValsAuth[2];
                    attrValsAuth[5] = attrVals[25] - attrValsAuth[2];

                    item.SetAttributeValue(attrAuth[4], attrValsAuth[4]);
                    item.SetAttributeValue(attrAuth[5], attrValsAuth[5]);

                    item.SetAttributeValue(attrAuth[6], GetFinanceVarianceLabelText(GetFinaceVarianceLabelIndex(attrValsAuth[4])));
                    item.SetAttributeValue(attrAuth[7], GetFinanceVarianceLabelText(GetFinaceVarianceLabelIndex(attrValsAuth[5])));


                    // more stuff
                    //item.SetAttributeValue(GetAttribute(story, "BFC11"), attrVals[21] - item.GetAttributeValueAsDouble(GetAttribute(story, "BFC09")));
                    //item.SetAttributeValue(GetAttribute(story, "BFC12"), attrVals[21] - item.GetAttributeValueAsDouble(GetAttribute(story, "BFC10")));

                }
                else
                {
                    LogError(string.Format("87 Could not find item {0}", itemName));
                }
            }
            catch (Exception ex)
            {
                LogError(string.Format("88 Could not find item {0}", ex.Message));
            }
        }

        private  SC.API.ComInterop.Models.Attribute GetAttribute(Story story, string id)
        {
            var attName = ConfigurationManager.AppSettings[id];
            var att = story.Attribute_FindByName(attName);
            if (att == null)
            {
                LogError(String.Format($"89 The lookupID was [{id}], but could not find attribute '{attName}' in story '{story.Name}'"));
            }

            return att;
        }

        private void ImportCosts(Story story, string projID)
        {
            try
            {
                string filelocation = $"{_workingFolder}/data/Costs/P{_period:D2}/{string.Format(ConfigurationManager.AppSettings["CostXLSFilename"], projID)}";
                var XLR = new Application();
                if (!File.Exists(filelocation))
                {
                    Log($"Cost Excel Doc is missing: " + filelocation);
                    return;
                }

                Log($"Opening Excel Doc " + filelocation);
                var wbR = XLR.Workbooks.Open(filelocation);

                var sheet = "Current Period PPR Cost";

                if (XLR.Sheets["Do Not Touch"].Cells(1, 7).Text != "SCAuthorised")
                {
                    LogError("Template is not a valid SharpCloud template.");
                    goto quit;
                }

                var fieldsToProcess = int.Parse(ConfigurationManager.AppSettings["costFieldCount"]);
                for (int rec = 1; rec <= fieldsToProcess; rec++)
                {
                    var data = ConfigurationManager.AppSettings[$"costField{rec:D4}"];
                    ProcessXLSRecord(story, XLR, data);
                }

                quit:
                // close Excel down
                wbR.Close(false);
                XLR.Quit();
                Marshal.ReleaseComObject(XLR);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
            }
        }

        private void ProcessXLSRecord(Story story, Application xls, string data)
        {
            var parts = data.Split('|');

            if (parts.Count() != 5)
            {
                LogError($"Congig setting does not consits of 5 parts '{data}'.");
                return;
            }
            try
            {
                var itemName = string.Format(parts[0], _period);
                var attributeName = parts[1];
                var sheetName = parts[2];
                var row = int.Parse(parts[3]);
                var col = int.Parse(parts[4]);

                // check item exists
                Item item = story.Item_FindByName(itemName);
                if (item == null)
                {
                    LogError($"Could not find an item called '{itemName}'.");
                    return;
                }
                Attribute attr = story.Attribute_FindByName(attributeName);
                if (attr == null)
                {
                    LogError($"Could not find attribute called '{attributeName}'.");
                    return;
                }

                SetAttributeWithLogging(item, attr, xls.Sheets[sheetName].Cells(row, col).Text);
            }
            catch (Exception e)
            {
                LogError($"Problem with line '{data}': '{e.Message}'.");
            }
        }


        private void ImportRisks(Story story, string projID)
        {
            try
            {
                string filelocation = $"{_workingFolder}/data/Risks/P{_period:D2}/{string.Format(ConfigurationManager.AppSettings["RiskXLSFilename"], projID)}";//, _period, projID);
                var XLR = new Application();
                if (!File.Exists(filelocation))
                {
                    Log($"Risk Excel Doc is missing: " + filelocation);
                    return;
                }

                Log($"Opening Excel Doc " + filelocation);
                var wbR = XLR.Workbooks.Open(filelocation);

                int row = 2;
                bool readline = true;
                var sheet = "Programme RISKS";

                int colName = 2;
                int colDesc = 3;
                int colOpenClose = 7;
                int colRiskScore = 14;

                int colMitigation = 15;
                int colDueDate = 16;
                int colPostRiskScore = 20;

                double riskMinScore = double.Parse(ConfigurationManager.AppSettings["RiskMinScore"]);

                var attrTitle = GetAttribute(story, "RiskAttributeTitle");
                var attrDesc = GetAttribute(story, "RiskAttributeDescription");
                var attrMit = GetAttribute(story, "RiskAttributeMitigation");
                var attrDue = GetAttribute(story, "RiskAttributeActionDue");
                var attrPreScore = GetAttribute(story, "RiskAttributePreMitScore");
                var attrPostScore = GetAttribute(story, "RiskAttributePostMitScore");

                var attrTimeRAG = GetAttribute(story, "RiskAttributeTimeRAG");
                var attrProjStatus = GetAttribute(story, "RiskAttributeProjectStatus");

                while (readline)
                {
                    string name = XLR.Sheets[sheet].Cells(row, colName).Text;
                    string open = XLR.Sheets[sheet].Cells(row, colOpenClose).Text.ToLower();
                    string score = XLR.Sheets[sheet].Cells(row, colRiskScore).Text;

                    if (!string.IsNullOrWhiteSpace(name))
                    { // keep reading until there is a blank row
                        var dScore = double.Parse(score);

                        if (open == "open" && dScore >= riskMinScore)
                        {
                            // we want this risk
                            string itemName = string.Format(ConfigurationManager.AppSettings["RiskName"], _period,
                                row - 1);

                            string desc = XLR.Sheets[sheet].Cells(row, colDesc).Text;
                            string mitigation = XLR.Sheets[sheet].Cells(row, colMitigation).Text;
                            string dueDate = XLR.Sheets[sheet].Cells(row, colDueDate).Text;
                            string postScore = XLR.Sheets[sheet].Cells(row, colPostRiskScore).Text;

                            var item = story.Item_FindByName(itemName);
                            if (item == null)
                                item = story.Item_AddNew(itemName, false);
                            item.Description = name;
                            item.Category = story.Category_FindByName(ConfigurationManager.AppSettings["RiskCategory"]);
                            item.Tag_AddNew(string.Format("P{0:D2}", _period));

                            item.SetAttributeValue(attrTitle, name);
                            item.SetAttributeValue(attrDesc, desc);
                            item.SetAttributeValue(attrMit, mitigation);
                            item.SetAttributeValue(attrDue, dueDate);
                            item.SetAttributeValue(attrPreScore, score);
                            item.SetAttributeValue(attrPostScore, postScore);

/*
                            if (!item.GetAttributeIsAssigned(attrTimeRAG))
                                item.SetAttributeValue(attrTimeRAG, GetVarianceLabel(-1));
                            if (!item.GetAttributeIsAssigned(attrProjStatus))
                                item.SetAttributeValue(attrProjStatus, GetVarianceLabel(-1));
*/
                            if (!item.GetAttributeIsAssigned(attrTimeRAG))
                            {
                                var dueDate2 = item.GetAttributeValueAsDate(attrDue);
                                if (dueDate2 != null)
                                {
                                    item.SetAttributeValue(attrTimeRAG,
                                        GetVarianceLabel(GetVarianceInt(DateTime.UtcNow - (DateTime) dueDate2)));
                                }
                                else
                                {
                                    item.SetAttributeValue(attrTimeRAG, GetVarianceLabel(-1));
                                }
                            }
                            else
                            {
                                Log($"Not setting time Time RAG on '{item.Name}' because value already set.");
                            }

                            if (!item.GetAttributeIsAssigned(attrProjStatus))
                            {
                                var score2 = item.GetAttributeValueAsDouble(attrPostScore);
                                item.SetAttributeValue(attrProjStatus, GetProjectVarianceLabelText(GetRiskLabelIndex(score2)));
                            }
                            else
                            {
                                Log($"Not setting time Project RAG on '{item.Name}' because value already set.");
                            }
                        }
                    }
                    else
                        readline = false;
                    row++;
                }
                // close Excel down
                wbR.Close(false);
                XLR.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(XLR);

                XLR = null;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
            }
        }

        private void SetAttributeWithLogging(Item item, SC.API.ComInterop.Models.Attribute att, int value)
        {
            SetAttributeWithLogging(item, att, $"{value:D2}");
        }

        private void SetAttributeWithLogging(Item item, SC.API.ComInterop.Models.Attribute att, string value)
        {
            try
            {
                Debug.WriteLine($"{att.Name} {value}");
                item.SetAttributeValue(att, value);
                Debug.WriteLine($"{att.Name} {item.GetAttributeValueAsText(att)}");
            }
            catch (Exception e)
            {
                LogError($"{e.Message}, value='{value}', item='{item.Name}', attribute='{att.Name}'");
            }
        }


        private void ImportBaseslineCosts(Application XL, Workbook wb, Story story, string projID)
        {
            if (wb == null)
            {
                return;
            }

            try
            {
                int attributeRowName = 2;
                int itemRowName = 3;
                int colProj = 1;
                int colStart = 2;

                int firstSheet = int.Parse(ConfigurationManager.AppSettings["BaselineCostsFirstSheet"]);
                int LastSheet = int.Parse(ConfigurationManager.AppSettings["BaselineCostsLastSheet"]);


                for (int sheet = firstSheet; sheet <= LastSheet; sheet++)
                {
                    int row = 5;
                    bool readline = true;
                    while (readline)
                    {
                        string s = XL.Sheets[sheet].Cells(row, colProj).Text;

                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            string projName = XL.Sheets[sheet].Cells(row, colProj).Text;
                            if (projName.Trim() == projID.Trim())
                            {
                                // we are now on the right row
                                bool readCol = true;
                                int col = 2;
                                while (readCol)
                                {
                                    var itemName = XL.Sheets[sheet].Cells(itemRowName, col).Text.Trim();

                                    if (!string.IsNullOrEmpty(itemName))
                                    {
                                        var attrName = XL.Sheets[sheet].Cells(attributeRowName, col).Text.Trim();
                                        var attribute = story.Attribute_FindByName(attrName);
                                        if (attribute == null)
                                        {
                                            LogError(string.Format($"90 We could not find the attribute '{attrName}' in story '{story.Name}'"));
                                        }
                                        else
                                        {
                                            var item = story.Item_FindByName(itemName);
                                            if (item == null)
                                            {
                                                LogError(string.Format($"91 We could not find the item '{itemName}' in story '{story.Name}'"));
                                            }
                                            else
                                            {
                                                var text = XL.Sheets[sheet].Cells(row, col).Text;
                                                item.SetAttributeValue(attribute, text);
                                            }
                                        }

                                    }
                                    else
                                        readCol = false;
                                    col++;
                                }
                            }
                        }
                        else
                            readline = false;
                        row++;
                    }
                }
            }
            catch (Exception e)
            {
                LogError("92 " + e.Message);
            }
        }

        private void LoadMilestones(Application XL, Workbook wb)
        {
            Log("Loading Milestone Data...");

            _milesStones = new List<MilestoneData>();

            if (wb == null)
            {
                return;
            }

            int row = 5;
            bool readline = true;
            int sheet = 2;

            int colProj = 1;
            int colName = 2;
            int colDesc = 3;
            int colBLFinish = 5;
            int colFinish = 7;
            int colPDM = 8;
            int colAPM = 9;

            int counter = 0;

            while (readline)
            {
                string s = XL.Sheets[sheet].Cells(row, 1).Text;

                if (!string.IsNullOrWhiteSpace(s))
                {
                    string pdm = XL.Sheets[sheet].Cells(row, colPDM).Text;
                    string apm = XL.Sheets[sheet].Cells(row, colAPM).Text;

                    bool bPDM = pdm.Trim().ToUpper() == "Y";
                    bool bAPM = apm.Trim().ToUpper() == "Y";

                    if (bPDM || bAPM)
                    {
                        string projName = XL.Sheets[sheet].Cells(row, colProj).Text.Trim();
                        string name = XL.Sheets[sheet].Cells(row, colName).Text.Trim();
                        string desc = XL.Sheets[sheet].Cells(row, colDesc).Text.Trim();
                        string blFinish = XL.Sheets[sheet].Cells(row, colBLFinish).Text.Trim();
                        string finish = XL.Sheets[sheet].Cells(row, colFinish).Text.Trim();

                        counter++;
                        _milesStones.Add(new MilestoneData
                        {
                            ProjId = projName,
                            ActivityId = name,
                            ActivityName = desc,
                            BLFinish = getDateString(blFinish),
                            Finish = getDateString(finish),
                            PDM = bPDM,
                            APM = bAPM
                        });
                    }
                }
                else
                    readline = false;
                row++;
            }

            Log($"Processed {counter} of {row} rows from milestone data");
        }


        private void ImportMilestones(Story story, string projID)
        {
            if (_milesStones == null)
                return; // no milestone data

            var attrBaseline = GetAttribute(story, "MilestoneBaselineFinish");
            var attrLast = GetAttribute(story, "MilestoneLastFinish");
            var attrForecast = GetAttribute(story, "MilestoneForecastFinish");
            var attrVarPrev = GetAttribute(story, "MilestoneForecastPreviousVariance");
            var attrVarBase = GetAttribute(story, "MilestoneForecastBaselineVariance");
            var attrTimeRAG = GetAttribute(story, "MilestoneForecastTimeRAG");
            var attrProjStatus = GetAttribute(story, "MilestoneForecastProjectStatus");

            int counter = 0;

            foreach (var msd in _milesStones.Where(m => m.ProjId == projID))
            {
                var desc = string.Format(ConfigurationManager.AppSettings["MilestoneName"], _period, msd.ActivityName);
                var extId = string.Format("{0}_{1}", msd.ActivityId, _period);
                var item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(desc, false);
                item.Name = desc;
                item.Category = story.Category_FindByName(ConfigurationManager.AppSettings["MilestoneCategory"]);
                item.ExternalId = extId;
                item.Tag_AddNew(string.Format("P{0:D2}", _period));

                item.SetAttributeValue(attrBaseline, msd.BLFinish);
                item.SetAttributeValue(attrForecast, msd.Finish);

                if (msd.PDM)
                {
                    item.Tag_AddNew(ConfigurationManager.AppSettings["MilestonePDMTag"]);
                }
                else if (msd.APM)
                {
                    item.Tag_AddNew(ConfigurationManager.AppSettings["MilestoneAPMTag"]);
                }

                var lastextId = string.Format("{0}_{1}", msd.ActivityId.Trim(), _period - 1);
                var itemLast = story.Item_FindByExternalId(lastextId);
                if (itemLast != null)
                {
                    // copy the last date
                    item.SetAttributeValue(attrLast, itemLast.GetAttributeValueAsText(attrForecast));
                }

                // calculate variances
                var baseline = item.GetAttributeValueAsDate(attrBaseline);
                var forecast = item.GetAttributeValueAsDate(attrForecast);
                var previous = item.GetAttributeValueAsDate(attrLast);

                var variance1 = forecast - baseline;
                var variance2 = forecast - previous;
                int max1 = -1, max2 = -1;
                if (variance1 != null)
                {
                    max1 = GetVarianceInt((TimeSpan)variance1);
                    item.SetAttributeValue(attrVarBase, GetVarianceLabel(max1));
                }
                if (variance2 != null)
                {
                    max2 = GetVarianceInt((TimeSpan)variance2);
                    item.SetAttributeValue(attrVarPrev, GetVarianceLabel(max2));
                }

                if (!item.GetAttributeIsAssigned(attrTimeRAG))
                    item.SetAttributeValue(attrTimeRAG, GetVarianceLabel(-1));
                else
                {
                    item.SetAttributeValue(attrTimeRAG, GetVarianceLabel((max1 > max2) ? max1 : max2));
                }

                if (!item.GetAttributeIsAssigned(attrProjStatus))
                    item.SetAttributeValue(attrProjStatus, GetVarianceLabel(-1));

            }
        }

        private static int GetVarianceInt(TimeSpan delay)
        {
            if (delay.Days > 28)
            {
                return 2;
            }
            else if (delay.Days > 0)
            {
                return 1;
            }
            return 0;
        }

        private static string GetVarianceLabel(int i)
        {
            switch (i)
            {
                case 2:
                    return ConfigurationManager.AppSettings["TimeVariance3"];//"> 4 wks late";
                case 1:
                    return ConfigurationManager.AppSettings["TimeVariance2"];//"0 - 4 wks late"; ;
                case 0:
                    return ConfigurationManager.AppSettings["TimeVariance1"];//"On time";
            }
            return ConfigurationManager.AppSettings["TimeVariance0"];//"No RAG assigned";
        }

        private static int GetBudgetVarianceLabelIndex(double v, double b)
        {
            if (b == 0 && v == 0)
                return 0; // special case - both are zero
            if (b == 0 && v != 0)
                return 2; // special case - prevent divide by zero

            if (v <= 0)
            {
                return 0; // less than 5%
            }
            else if (v / b <= 0.1)
            {
                return 1; // 5-10%
            }
            return 2; //"more than 10%";
        }

        private static string GetBudgetVarianceLabelText(int i)
        {
            switch (i)
            {
                case 2:
                    return ConfigurationManager.AppSettings["BudgetVariance3"];//"more than 10%";
                case 1:
                    return ConfigurationManager.AppSettings["BudgetVariance2"];//"5% to 10%";
                case 0:
                    return ConfigurationManager.AppSettings["BudgetVariance1"];//"less than 5%";
            }
            return ConfigurationManager.AppSettings["BudgetVariance0"];//"No RAG assigned";
        }


        private static int GetCostVarianceLabelIndex(double v, double b)
        {
            if (b == 0 && v == 0)
                return 0; // special case - both are zero
            if (b == 0 && v != 0)
                return 2; // special case - prevent divide by zero

            if (Math.Abs(v / b) < 0.05)
            {
                return 0; // less than 5%
            }
            if (Math.Abs(v / b) <= 0.1)
            {
                return 1; // 5-10%
            }
            return 2; //"more than 10%";
        }

        private static string GetCostVarianceLabelText(int i)
        {
            switch (i)
            {
                case 2:
                    return ConfigurationManager.AppSettings["CostVariance3"];//"more than 10%";
                case 1:
                    return ConfigurationManager.AppSettings["CostVariance2"];//"5% to 10%";
                case 0:
                    return ConfigurationManager.AppSettings["CostVariance1"];//"less than 5%";
            }
            return ConfigurationManager.AppSettings["CostVariance0"];//"No RAG assigned";
        }

        private static string GetProjectVarianceLabelText(int i)
        {
            switch (i)
            {
                case 2:
                    return ConfigurationManager.AppSettings["ProjectVariance3"];//Action Required
                case 1:
                    return ConfigurationManager.AppSettings["ProjectVariance2"];//Monitoring
                case 0:
                    return ConfigurationManager.AppSettings["ProjectVariance1"];//Okay;
            }
            return ConfigurationManager.AppSettings["ProjectVariance0"];//No RAG assigned
        }

        private static int GetRiskLabelIndex(double score)
        {
            if (score <= 0)
                return -1;
            if (score <= 10)
                return 0;
            if (score <= 18)
                return 1;

            return 2; // very bad
        }

        private static int GetFinaceVarianceLabelIndex(double v)
        {
            if (v > 0)
            {
                return 1; // over budget
            }
            return 0; // on or under budget
        }

        private static string GetFinanceVarianceLabelText(int i)
        {
            switch (i)
            {
                case 1:
                    return ConfigurationManager.AppSettings["FinanceVariance2"]; // over budget
                case 0:
                    return ConfigurationManager.AppSettings["FinanceVariance1"]; // on or under budget
            }
            return ConfigurationManager.AppSettings["CostVariance0"];//"No RAG assigned";
        }


        private string getProjectID(string str)
        {
            return str.TrimStart().Substring(0, 9).Trim();
        }

        private static string getDateString(string strDate)
        {
            return strDate.Replace(" A", "").Replace("*", "").Trim();
        }


        private void GetStoryPanels()
        {
            var userid = ConfigurationManager.AppSettings["userid"];
            var passwd = ConfigurationManager.AppSettings["passwd"];

            var sc = new SharpCloudApi(userid, passwd);

            var story = sc.LoadStory("188421be-07f2-45ef-b392-fa512ffd19d7");

            foreach (var i in story.Items)
            {
                foreach (var p in i.Panels)
                {
                    if (p.Type == Panel.PanelType.Image && p.Data != "[]")
                    {
                        Log($"--------------------------------");
                        Log($"Category = {i.Category.Name}");
                        Log($"Name = {i.Name}");
                        Log($"Title = {p.Title}");
                        Log($"Data = {p.Data}");
                    }
                }
            }
        }

        private void SetStoryPanels()
        {
            var userid = ConfigurationManager.AppSettings["userid"];
            var passwd = ConfigurationManager.AppSettings["passwd"];

            var sc = new SharpCloudApi(userid, passwd);

            var story = sc.LoadStory("188421be-07f2-45ef-b392-fa512ffd19d7");
            SetStoryPanels(story);

            story.Save();
        }
        private void SetStoryPanels(Story story)
        {
            foreach (var i in story.Items)
            {
                SetPanelData(i);
            }
        }

        private void SetPanelData(Item item)
        {
            var data = "";
            var title = "";
            var data2 = "";
            var title2 = "";

            switch (item.Category.Name.Trim())
            {
                case "PRG Project Details":
                    title = "Step 1: Project Details";
                    data = "[\"9ea127a8-e076-4154-a1de-c3221c042c24\"]";
                    break;
                case "Health Safety and Environmental":
                    title = "Step 2: HSE (Health Safety and Environmental";
                    data = "[\"144bca58-a72d-4562-8baf-fc0238c6abb9\"]";
                    break;
                case "Project Milestones":
                    title = "Step 3: Project Milestones";
                    data = "[\"53243b8d-29d3-4a18-b0c8-96bbc6cbc0b9\",\"e5798bfb-1463-4707-b4e0-e9f23133a7d3\"]";
                    break;
                case "Pending Changes":
                    title = "Step 4: Pending Changes";
                    data = "[\"f63cf887-6af5-423d-a993-68e550198ec3\",\"9790f2f6-773b-4b5d-b320-7e8f775e5bcd\"]";
                    break;
                case "Risk":
                    title = "Step 5: Risks";
                    data = "[\"e1aafa9d-f97f-4a6d-8b3a-948872f20ddc\",\"c2f3bb51-9e0c-4e63-bfe3-9ea3a98bccf3\"]";
                    break;
                case "Financial":
                    title = "Step 6: Budget Forecast Cost (BFC) and Estimated Final Cost (EFC)";
                    data = "[\"983830b8-1e1a-4e7d-b6e2-66f09f4e6b61\",\"63369def-9f6d-4e67-aa90-2d89063baa6e\",\"a1f36915-e917-417e-b0b3-6b5cdc7e1099\",\"4f64fcd2-1ce4-4714-9d8c-1899d0821898\"]";
                    break;
                case "Project Manager's Commentary":
                    title = "Step 7: Project Managers Commentary";
                    data = "[\"bc3e7ab1-505a-4347-bab1-ee8c6a119921\"]";
                    title2 = "Step 8: Set Overall Project Status";
                    data2 = "[\"670c796c-e05d-457d-9ebf-d91cd39df495\"]";
                    break;
                case "Report Approval":
                    title = "Step 9: Review and Approve Project";
                    data = "[\"e096e675-9976-4bac-887b-9f69803d5098\"]";
                    break;
            }

            if (data != "")
            {
                var panel = item.Panel_FindByTitle(title) ?? item.Panel_Add(title, Panel.PanelType.Image);
                panel.Data = data;
            }
            if (data2 != "")
            {
                var panel = item.Panel_FindByTitle(title2) ?? item.Panel_Add(title2, Panel.PanelType.Image);
                panel.Data = data2;
            }
        }

        public class MilestoneData
        {
            public string ProjId;
            public string ActivityId;
            public string ActivityName;
            public string BLFinish;
            public string Finish;
            public bool PDM;
            public bool APM;
        }
    }
}
