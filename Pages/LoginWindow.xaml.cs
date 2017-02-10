using Uniconta.Common;
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
using System.Windows.Shapes;
using ImportingTool.Utility;

namespace ImportingTool.Pages
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        static Guid ImportToolGuid = new Guid("00000000-0000-0000-0000-000000000000");

        public LoginWindow()
        {
            InitializeComponent();
            this.loginCtrl.loginButton.Click += loginButton_Click;
            this.loginCtrl.CancelButton.Click += CancelButton_Click;
        }

        void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        async void loginButton_Click(object sender, RoutedEventArgs e)
        {
            string password = loginCtrl.Password;
            string userName = loginCtrl.UserName;

            if (ImportToolGuid == Guid.Empty)
            {
                MessageBox.Show("You need to have your own valid App id, obtained at your Uniconta partner or at Uniconta directly", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ErrorCodes errorCode = await SetLogin(userName, password);

            if (errorCode == ErrorCodes.Succes)
            {
                Uniconta.ClientTools.Localization.SetDefault((Language)SessionInitializer.CurrentSession.User._Language);
                this.DialogResult = true;
            }
            else
                UtilFunctions.ShowErrorMessage(errorCode);
        }

        private async Task<ErrorCodes> SetLogin(string username, string password)
        {
            if (! ValidateUserCredentials(username, password))
                return ErrorCodes.NoSucces;

            try
            {
                var ses = SessionInitializer.GetSession();
                return await ses.LoginAsync(username, password, Uniconta.Common.User.LoginType.PC_Windows, ImportToolGuid);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
                return ErrorCodes.Exception;
            }
        }

        private bool ValidateUserCredentials(string username, string password)
        {
            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                return true;

            return false;
        }
    }
}
