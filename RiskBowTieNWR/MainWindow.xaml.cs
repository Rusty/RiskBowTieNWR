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
    }
}
