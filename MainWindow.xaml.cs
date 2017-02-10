using ImportingTool.Model;
using ImportingTool.Pages;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Uniconta.API.Service;
using Uniconta.API.System;
using Uniconta.Common;
using Uniconta.DataModel;

namespace ImportingTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool _terminate;
        public bool Terminate { get { return _terminate; } set { _terminate = value; BreakConversion(); } }
        public MainWindow()
        {
            this.DataContext = this;
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
            this.MouseDown += MainWindow_MouseDown;
            btnCopyLog.Content = String.Format(Uniconta.ClientTools.Localization.lookup("CopyOBJ"), Uniconta.ClientTools.Localization.lookup("Logs"));
            progressBar.Visibility = Visibility.Collapsed;

            /* Remove it*/
            progressBar.Visibility = Visibility.Visible;
            progressBar.Maximum = 100;
            UpdateProgressbar(0, null);
        }

        private void MainWindow_MouseDown(object sender, MouseButtonEventArgs e)
        {
            btnImport.IsEnabled = true;
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            DisplayLoginScreen();
            cmbImportFrom.ItemsSource = Enum.GetNames(typeof(ImportFrom));
            cmbImportDimension.ItemsSource = new List<string>() { "Ingen", "Kun Afdeling", "Afdeling, Bærer", "Afdeling, Bærer, Formål" };
            cmbImportDimension.SelectedIndex = 3;
        }
        async private void DisplayLoginScreen()
        {
            var logon = new LoginWindow();
            logon.Owner = this;
            logon.ShowDialog();

            if (logon.DialogResult.HasValue && logon.DialogResult.Value)
            {
                await SessionInitializer.SetupCompanies();
            }
            else
                this.Close();
        }

        public async void Import(importLog log)
        {
            UpdateProgressbar(0, "Starting..");

            var ses = SessionInitializer.CurrentSession;

            Company cc = new Company();
            cc._Name = txtCompany.Text;
            cc._CurrencyId = Currencies.DKK;
            var country = CountryCode.Denmark;

            if (cmbImportFrom.SelectedIndex <= 1)
            {
                // c5
                cc._ConvertedFrom = (int)ConvertFromType.C5;
                log.extension = ".def";
                log.ConvertDanishLetters = true;
                log.CharBuf = new byte[1];
                if (cmbImportFrom.SelectedIndex == 0)
                    log.C5Encoding = Encoding.GetEncoding(850); // western Europe. 865 is Nordic
                else
                {
                    country = CountryCode.Iceland;
                    cc._CurrencyId = Currencies.ISK;
                    log.C5Encoding = Encoding.GetEncoding(861); // Iceland
                }

                if (!log.OpenFile("exp00000", "Definitions"))
                {
                    log.AppendLogLine("These is no c5-files in this directory");
                    return;
                }
                log.extension = ".kom";
            }
            else
            {
                // eco
                cc._ConvertedFrom = (int)ConvertFromType.Eco;
                if (cmbImportFrom.SelectedIndex == 3) // Norsk Eco
                {
                    country = CountryCode.Norway;
                    cc._CurrencyId = Currencies.NOK;
                }

                log.extension = ".csv";
                log.FirstLineIsEmpty = true;

                if (!log.OpenFile(country == CountryCode.Denmark ? "RegnskabsAar" : "Regnskapsaar"))
                {
                    log.AppendLogLine("Der er ingen filer i dette bibliotek fra e-conomic");
                    return;
                }
            }

            var lin = log.ReadLine();
            if (lin == null)
            {
                log.AppendLogLine("Filen er tom");
                log.closeFile();
            }

            char splitCh, apostrof;
            var cols = lin.Split(';');
            if (cols.Length > 1)
                splitCh = ';';
            else
            {
                cols = lin.Split(',');
                if (cols.Length > 1)
                    splitCh = ',';
                else
                    splitCh = '#';
            }

            if (lin[0] == '\'')
                apostrof = '\'';
            else
                apostrof = '"';
            log.sp = new StringSplit(splitCh, apostrof);

            cc._CountryId = country;

            UpdateProgressbar(0, "Opret firma..");

            var err = await ses.CreateCompany(cc);
            if (err != 0)
            {
                log.AppendLogLine("Cannot create Company !");
                return;
            }

            Company fromCompany = null;
            var ccList = await ses.GetStandardCompanies(country);
            foreach (var c in ccList)
                if ((c.CompanyId == 19 && country == CountryCode.Denmark) || // 19 = Dev/erp
                    (c.CompanyId == 1855 && country == CountryCode.Iceland) || // 1855 = erp
                    (c.CompanyId == 108 && country == CountryCode.Norway)) // 108 = Dev/erp
                {
                    fromCompany = c;
                    break;
                }
            if (fromCompany == null)
            {
                log.AppendLogLine("Cannot find standard Company !");
                return;
            }

            CompanyAPI comApi = new CompanyAPI(ses, cc);
            err = await comApi.CopyBaseData(fromCompany, cc, false, true, true, false, true, false, false, false);
            if (err != 0)
            {
                log.AppendLogLine(string.Format("Copy of default company failed. Error = {0}", err.ToString()));
                log.AppendLogLine("Conversion has stopped");
                return;
            }

            log.AppendLogLine("Import started....");

            log.api = new CrudAPI(ses, cc);

            log.CompCurCode = log.api.CompanyEntity._CurrencyId;
            log.CompCur = log.CompCurCode.ToString();
            log.CompCountryCode = log.api.CompanyEntity._CountryId;
            log.CompCountry = log.CompCountryCode.ToString();

            if (cmbImportFrom.SelectedIndex <= 1)
            {
                err = await c5.importAll(log, cmbImportDimension.SelectedIndex);
            }
            else
            {
                err = await eco.importAll(log, cmbImportFrom.SelectedIndex - 1);
            }

            cc.InvPrice = (log.HasPriceList);
            cc.Inventory = (log.HasItem);
            cc.InvBOM = (log.HasBOM);
            cc.Creditor = (log.HasCreditor);
            cc.Debtor = (log.HasDebitor);
            cc.Contacts = (log.HasContact);
            cc.InvClientName = false;
            await log.api.Update(cc);

            log.closeFile();

            log.AppendLogLine("Færdig !");
            if (err == 0)
            {
                log.AppendLogLine("");
                log.AppendLogLine("Du kan nu logge ind i Uniconta og åbne dit firma.");
                log.AppendLogLine("Der kan godt gå lidt tid inden alle dine poster er bogført.");
                log.AppendLogLine("Du kan slette firmaet i menuen Firma/Firmaoplysninger,");
                log.AppendLogLine("hvis du ønsker at foretage en ny import.");
                log.AppendLogLine("Held og lykke med dit firma i Uniconta.");
            }
            else
                log.AppendLogLine(string.Format("Conversion has stopped due to an error {0}", err.ToString()));
        }

        private void cmbCompanies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        FolderBrowserDialog openFolderDialog;
        private void btnImportFromDir_Click(object sender, RoutedEventArgs e)
        {
            openFolderDialog = new FolderBrowserDialog();
            if (openFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtImportFromDirectory.Text = openFolderDialog.SelectedPath;
            }
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtCompany.Text))
            {
                System.Windows.MessageBox.Show(String.Format(Uniconta.ClientTools.Localization.lookup("CannotBeBlank"), Uniconta.ClientTools.Localization.lookup("CompanyName")));
                return;
            }
            if (cmbImportFrom.SelectedIndex == -1)
            {
                System.Windows.MessageBox.Show("Please select C5 or e-conomic.");
                return;
            }
            if (cmbImportDimension.SelectedIndex > -1)
            {
               
            }
            var path = txtImportFromDirectory.Text;
            if (string.IsNullOrEmpty(path))
            {
                System.Windows.MessageBox.Show("Select directory");
                return;
            }

            txtCompany.Focus();
            btnImport.IsEnabled = false;

            if (path[path.Length - 1] != '\\')
                path += '\\';

            importLog log = new importLog(path);
            log.window = this;
            if (cmbImportFrom.SelectedIndex <= 1) // c5
                log.Set0InAccount = chkSetAccount.IsChecked.Value;
            log.target = txtLogs;
            txtLogs.DataContext = log;
            log.AppendLogLine(string.Format("Reading from path {0}", path));
            Import(log);
        }

        private void cmbImportFrom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbImportFrom.SelectedIndex <= 1) // c5
            {
                chkSetAccount.IsEnabled = true;
                cmbImportDimension.IsEnabled = true;
            }
            else
            {
                chkSetAccount.IsEnabled = false;
                cmbImportDimension.IsEnabled = false;
            }
        }

        private void btnCopyLog_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(txtLogs.Text);
        }

        private void btnTerminate_Click(object sender, RoutedEventArgs e)
        {
            var res = System.Windows.MessageBox.Show("Are you sure?", Uniconta.ClientTools.Localization.lookup("Confirmation"), MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                this.Terminate = true;
            }
        }

        void BreakConversion()
        {
        }

        public void UpdateProgressbar(int value, string text)
        {
            progressBar.Value = value;
            if (!string.IsNullOrEmpty(text))
                progressBar.Tag = text;
            else
                progressBar.Tag = string.Format("{0} of {1}", value, progressBar.Maximum);
        }
    }
    public class HeightConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (double)value - 30;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (double)value + 30;
        }
    }
    /* For ProgressBar*/
    public class RectConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            double width = (double)values[0];
            double height = (double)values[1];
            return new Rect(0, 0, width, height);
        }
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public enum ImportFrom
    {
        c5_Danmark,
        c5_Iceland,
        economic_Danmark,
        economic_Norge,
    }
}
