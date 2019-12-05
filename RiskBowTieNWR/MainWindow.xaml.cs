using System;
using System.Collections.Generic;
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
using RiskBowTieNWR.Helpers;
using RiskBowTieNWR.ViewModels;
using RiskBowTieNWR.Views;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;

namespace RiskBowTieNWR
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

            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _viewModel.SaveAllData();
        }

        private void Hyperlink_OnClick(object sender, RoutedEventArgs e)
        {
            var proxy = new ProxySettings(_viewModel);
            proxy.ShowDialog();
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
                MessageBox.Show("Awesome! Your credentials appear to be correct.");
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

        private SharpCloudApi GetApi()
        {
            _viewModel.UserName = tbUsername.Text;
            _viewModel.Url = tbUrl.Text;
            _viewModel.Password = tbPassword.Password;

            return new SharpCloudApi(_viewModel.UserName, _viewModel.Password, _viewModel.Url, _viewModel.Proxy, _viewModel.ProxyAnnonymous, _viewModel.ProxyUserName, _viewModel.ProxyPassword);
        }


        private void SelectTeam_Click(object sender, RoutedEventArgs e)
        {
            var sel = new SelectTeam(GetApi());
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedTeam = sel.SelectedTeam;
            }
        }

        private void SelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sel = new SelectStory(GetApi(), false);
                if (sel.ShowDialog() == true)
                {
                    _viewModel.SelectedTemplateStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
                    CheckTemplate();
                }
            }
            catch (Exception E)
            {
                
            }
        }

        private async void CheckTemplate()
        {
            var log = new Logger(_viewModel, 2);
            log.Log($"Checking template {_viewModel.SelectedTemplateStory.Id}");
            await Task.Delay(100);
            var sc = GetApi();
            var template = sc.LoadStory(_viewModel.SelectedTemplateStory.Id);
            RiskModel.EnsureStoryHasRightStructure(template, log);
            template.Save();
            _viewModel.ShowWaitForm = false;
        }

        private void SelectPortfolio_Click(object sender, RoutedEventArgs e)
        {
            var sel = new SelectStory(GetApi(), false);
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedPortfolioStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
            }
        }

        private void SelectControl_Click(object sender, RoutedEventArgs e)
        {
            var sel = new SelectStory(GetApi(), false);
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedControlStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
            }
        }


        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                _viewModel.SelectedDataFolder = dialog.SelectedPath;
            }
        }

        private void mainTab_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _viewModel.LoadFileList();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _viewModel.FileList)
                f.IsSelected = true;

            listFiles.ItemsSource = null;
            listFiles.ItemsSource = _viewModel.FileList;
        }

        private void SelectNon_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _viewModel.FileList)
                f.IsSelected = false;

            listFiles.ItemsSource = null;
            listFiles.ItemsSource = _viewModel.FileList;
        }

        private void ProcessFiles_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.FileList.Where(f => f.IsSelected==true).Any())
            {
                MessageBox.Show("Please make sure you select some files from the list on the left");
                return;
            }

            ProcessSelectedFiles();
        }

       private void SelectMigrate_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.FileList.Where(f => f.IsSelected==true).Any())
            {
                MessageBox.Show("Please make sure you select some files from the list on the left");
                return;
            }

            MigrateSelectedFiles();
        }
        private void ProcessPortfolio_Click(object sender, RoutedEventArgs e)
        {
            ProcessPortfolio();
        }

        private async void ProcessPortfolio()
        {
            _viewModel.ProgressLogText2 = ""; // clear
            var logger = new Logger(_viewModel, 2);

            // find list of stories for selected team
            logger.Log($"Reading Existing Stories in {_viewModel.SelectedTeamName}...");
            await Task.Delay(100);
            var client = GetApi();

            RiskModel.ProcessBowTies(client, _viewModel.SelectedTeam.Id, _viewModel.SelectedPortfolioStory.Id,
                _viewModel.SelectedControlStory.Id, _viewModel.SelectedTemplateStory.Id, logger);
           
        }

        private async void MigrateSelectedFiles()
        {
            _viewModel.ProgressLogText = ""; // clear
            var logger = new Logger(_viewModel);

            System.IO.Directory.CreateDirectory(_viewModel.SelectedDataFolder + "/migrated/");

            try
            {
                // find list of stories for selected team
                foreach (var f in _viewModel.FileList.Where(fi => fi.IsSelected == true))
                {
                    if (f.FileName[0] == '~')
                    {
                        logger.Log($"Skipping {f.FileName}");
                        await Task.Delay(100);
                        continue;
                    }

                    logger.Log($"Processing {f.FileName}");
                    await Task.Delay(100);

                    RiskModel.MigrateSpreadsheet(f.FullPath, _viewModel.SelectedDataFolder + "/Template.xltm",
                        _viewModel.SelectedDataFolder + "/migrated/" + f.FileName, logger);
                }
            }
            catch (Exception eBad)
            {
                logger.LogError(eBad.Message);
            }

            await Task.Delay(1000);

            _viewModel.ShowWaitForm = false;
        }

        private async void ProcessSelectedFiles()
        {
            _viewModel.ProgressLogText = ""; // clear
            var logger = new Logger(_viewModel, 1);

            // find list of stories for selected team
            logger.Log($"Reading Existing Stories in {_viewModel.SelectedTeamName}...");
            await Task.Delay(100);
            var client = GetApi();
            var teamStories = client.StoriesTeam(_viewModel.SelectedTeam.Id);
            logger.Log($"{teamStories.Count()} stories loaded.");

            foreach (var f in _viewModel.FileList.Where(fi => fi.IsSelected == true))
            {
                if (f.FileName[0] == '~')
                {
                    logger.Log($"Skipping {f.FileName}");
                    await Task.Delay(100);
                    continue;
                }

                logger.Log($"Processing {f.FileName}");
                await Task.Delay(100);

                // does the risk story already exist?

                Story sampleStory = null;
                try
                {
                    logger.Log($"Loading sample '{_viewModel.SelectedTemplateStory.Name}'");
                    sampleStory = client.LoadStory(_viewModel.SelectedTemplateStory.Id);
                }
                catch (Exception e)
                {
                    logger.LogError($"Unable to load sample '{_viewModel.SelectedTemplateStory.Name}'");
                }

                if (sampleStory == null)
                {
                    logger.Log($"Unable to load the Sample '{_viewModel.SelectedTemplateStory.Name}' - PROCESS ABORTING!");
                    await Task.Delay(100);
                    return;
                }

                string storyId = RiskModel.GetExcelTemplateStoryID(f.FullPath, logger);
                if (string.IsNullOrEmpty(storyId))
                {
                    // does not exist so create
                    logger.Log($"Creating {f.Name} from Template '{_viewModel.SelectedTemplateName}'");
                    await Task.Delay(100);

                    var s = client.NewStory(f.Name, _viewModel.SelectedTemplateStory.Id);
                    if (s != null && s.Id != null)
                    {
                        try
                        {
                            s = client.LoadStory(s.Id);
                            s.StoryAsRoadmap.PackID = sampleStory.StoryAsRoadmap.PackID; // make sure new stories are part of the same pack
                            s.StoryAsRoadmap.TeamID = _viewModel.SelectedTeam.Id;
                            s.StoryAsRoadmap.ImageID = new Guid(_viewModel.SelectedTemplateStory.ImageId);
                            s.Save();
                            storyId = s.Id;
                        }
                        catch (Exception e1)
                        {
                            logger.LogError($"Unable to create story from file '{f.Name}' - {e1.Message}");
                        }

                        try
                        {
                            RiskModel.SetExcelTemplateStoryID(storyId, f.FullPath, logger);
                        }
                        catch (Exception e2)
                        {
                            logger.LogError($"Unable to update excel storyID '{storyId}' - {e2.Message}");
                        }

                        logger.Log($"{f.FileName} created '{_viewModel.SelectedTemplateName}'");
                    }
                    else
                    {
                        logger.LogError($"{f.FileName} was not created created from '{_viewModel.SelectedTemplateName}'");
                    }

                    await Task.Delay(100);
                }

                logger.Log($"Loading Control Story '{_viewModel.SelectedControlStory.Name}'");
                await Task.Delay(100);
                Story controlStory = null;
                try
                {
                    controlStory = client.LoadStory(_viewModel.SelectedControlStory.Id);
                    //RiskModel.EnsureStoryHasRightStructure(controlStory, logger);
                    //controlStory.Save();
                }
                catch (Exception ex)
                {
                    logger.LogError($"Unable to load story '{_viewModel.SelectedControlStory.Name}'");
                    await Task.Delay(100);
                }
                
                // now ready to load story and updates
                if (!string.IsNullOrEmpty(storyId))
                {
                    logger.Log($"Loading story id '{storyId}'");
                    await Task.Delay(100);
                    try
                    {
                        var story = client.LoadStory(storyId);
                        var version = "5";
                            if (chkVersion4.IsChecked== true) version = "4";
                        var sharepermission = story.StoryAsRoadmap.SharedUsers.FirstOrDefault(su => su.User.Username.ToLower() == _viewModel.UserName.ToLower()).Action.ToString();
                        if (sharepermission != null && (sharepermission == "admin" || sharepermission == "owner"))
                        {
                            RiskModel.CreateStoryFromXLTemplate(story, controlStory, f.FullPath, logger, chkDelete.IsChecked == true, chkDeleteRels.IsChecked == true, chkVerbose.IsChecked == true,version);
                            story.Save();
                        }
                        else
                        {
                            logger.LogError($"Skipping story '{story.Name}', as you only have '{sharepermission}' permission");
                            await Task.Delay(100);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Unable to load story '{storyId}'");
                        await Task.Delay(100);
                    }
                }

                controlStory.Save(); // save any changes to control library (nneded to save relationship info)
                controlStory = client.LoadStory(_viewModel.SelectedControlStory.Id);
                RiskModel.UpdateRiskCountOnControlStory(controlStory, logger);
                controlStory.Save(); // save any changes to control library (nneded to save relationship info)

            }

            await Task.Delay(1000);

            _viewModel.ShowWaitForm = false;
        }


        private void ViewLog_Click(object sender, RoutedEventArgs e)
        {
            Logger.ShowLogFolder();
        }
    }
}
