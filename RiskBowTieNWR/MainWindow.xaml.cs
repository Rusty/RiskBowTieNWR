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
            var sel = new SelectStory(GetApi(), false);
            if (sel.ShowDialog() == true)
            {
                _viewModel.SelectedTemplateStory = new Models.StoryLite2(sel.SelectedStoryLites[0]);
            }
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

        private async void ProcessSelectedFiles()
        {
            _viewModel.ProgressLogText = ""; // clear
            var logger = new Logger(_viewModel);

            // find list of stories for selected team
            logger.Log($"Reading Existing Stories in {_viewModel.SelectedTeamName}...");
            await Task.Delay(100);
            var client = GetApi();
            var teamStories = client.StoriesTeam(_viewModel.SelectedTeam.Id);
            logger.Log($"{teamStories.Count()} stories loaded.");

            foreach (var f in _viewModel.FileList.Where(fi => fi.IsSelected == true))
            {
                logger.Log($"Processing {f.FileName}");
                await Task.Delay(100);

                // does the risk story already exist?
                var ts = teamStories.FirstOrDefault(s => s.Name == f.Name);

                string storyId = null;
                if (ts == null)
                {
                    // does not exist so create
                    logger.Log($"Creating {f.Name} from Template '{_viewModel.SelectedTemplateName}'");
                    await Task.Delay(100);

                    var s = client.NewStory(f.Name, _viewModel.SelectedTemplateStory.Id);
                    if (s != null)
                    {
                        s.StoryAsRoadmap.TeamID = _viewModel.SelectedTeam.Id;
                        s.StoryAsRoadmap.ImageID = new Guid(_viewModel.SelectedTemplateStory.ImageId);
                        s.Save();

                        storyId = s.Id;
                        logger.Log($"{f.FileName} created '{_viewModel.SelectedTemplateName}'");
                    }
                    else
                    {
                        logger.LogError($"{f.FileName} was not created created from '{_viewModel.SelectedTemplateName}'");
                    }

                    await Task.Delay(100);
                }
                else
                {
                    storyId = ts.Id;
                }

                // now ready to load story and update
                if (!string.IsNullOrEmpty(storyId))
                {
                    logger.Log($"Loading {f.Name}''");
                    await Task.Delay(100);
                    var story = client.LoadStory(storyId);

                    RiskModel.CreateStoryFromXLTemplate(story, f.FullPath, logger);
                    story.Save();
                }


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
