using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Uniconta.API.System;
using Uniconta.ClientTools;
using Uniconta.Common;
using Uniconta.DataModel;

namespace ImportingTool.Model
{
    public class GLPostingLineLocal : Uniconta.API.GeneralLedger.GLPostingLine
    {
        public string OrgText;
        public int primo;
    }

    public class MyGLAccount : GLAccount
    {
        public bool HasVat, HasChanges;
    }

    public partial class importLog : INotifyPropertyChanged
    {
        public TextBox target;
        public string LogMsg { get { return LogMessage.ToString(); } set { NotifyPropertyChanged("LogMsg"); } }
        StringBuilder LogMessage = new StringBuilder();
        public void AppendLog(string msg)
        {
            LogMessage.Append(msg);
            LogMsg = LogMessage.ToString();
        }
        public void AppendLogLine(string msg)
        {
            LogMessage.AppendLine(msg);
            LogMsg = LogMessage.ToString();
            if (target != null)
                target.ScrollToEnd();
        }
        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public MainWindow window;
        public string extension;
        public string path;
        public StringSplit sp;
        public string CurFile;
        public CrudAPI api;
        public List<PaymentTerm> pTerms;
        TextReader r;
        public bool FirstLineIsEmpty, ConvertDanishLetters;
        public bool HasDepartment, HasPorpose, HasCentre, HasBOM, HasItem, HasPriceList, HasDebitor, HasCreditor, HasContact;
        public int NumberOfDimensions;
        public SQLCache LedgerAccounts, vats, Debtors, Creditors, PriceLists, Items, ItemGroups, Employees, Payments, DebGroups, CreGroups;
        public SQLCache dim1, dim2, dim3;
        public string errorAccount, valutaDif;
        public Encoding C5Encoding;
        public byte[] CharBuf;

        public bool VatIsUpdated;
        public Currencies CompCurCode;
        public string CompCur;
        public CountryCode CompCountryCode;
        public string CompCountry;

        public CompanyFinanceYear[] years;
        public int Accountlen;
        public bool Set0InAccount;

        public void closeFile()
        {
            if (r != null)
            {
                var _r = r;
                r = null;
                _r.Dispose();
            }
        }

        public importLog(string path)
        {
            this.path = path;
        }

        public void Ex(Exception ex)
        {
            api.ReportException(ex, string.Format("File={0}", CurFile));
            AppendLogLine(ex.Message);
        }

        public bool OpenFile(string file, string comment = null)
        {
            closeFile();

            if (comment != null)
                comment = string.Format("{0} - {1}", file, comment);
            else
                comment = file;

            CurFile = this.path + file + extension;
            if (!File.Exists(CurFile))
            {
                AppendLogLine(String.Format(Uniconta.ClientTools.Localization.lookup("FileNotFound"), comment));
                return false;
            }
            AppendLogLine(String.Format(Uniconta.ClientTools.Localization.lookup("ImportingFile"), comment));
            r = new StreamReader(CurFile, Encoding.Default, true);
            if (r != null)
            {
                if (FirstLineIsEmpty)
                    ReadLine();  // first line is empty
                return true;
            }
            AppendLogLine(string.Format("{0} : ", Uniconta.ClientTools.Localization.lookup("FileAccess"), CurFile));
            return false;
        }
        public void Log(string s)
        {
            AppendLogLine(s);
        }
        public async Task Error(ErrorCodes Err)
        {
            var errors = await api.session.GetErrors();
            var s = Err.ToString();
            api.ReportException(null, string.Format("Error in conversion = {0}, file={1}", s, CurFile));
            AppendLogLine(s);
            if (errors != null)
            {
                foreach(var er in errors)
                {
                    if (er.inProperty != null)
                        AppendLogLine(string.Format("In property = {0}", er.inProperty));
                    if (er.message != null)
                        AppendLogLine(string.Format("Message = {0}", er.message));
                }
            }
        }

        public string ReadLine() { return r.ReadLine(); }

        public List<string> GetLine(int minCols)
        {
            string oldLin = null;
            for (;;)
            {
                var lin = r.ReadLine();
                if (lin == null)
                    return null;
                if (ConvertDanishLetters)
                {
                    int StartIndex = 0;
                    int pos;
                    while ((pos = lin.IndexOf('\\', StartIndex)) >= 0)
                    {
                        StartIndex = pos + 1;
                        int val = 0;
                        for (int i = 1; i <= 3; i++)
                        {
                            var ch = lin[pos + i];
                            if (ch >= '0' && ch <= '7')
                                val = val * 8 + (ch - '0');
                            else
                                break;
                        }
                        if (val >= 128 && val <= 255)
                        {
                            lin = lin.Remove(pos, 4);
                            this.CharBuf[0] = (byte)val;
                            var ch = this.C5Encoding.GetString(this.CharBuf);
                            lin = lin.Insert(pos, ch);
                        }
                    }
                    /*
                    //******************************************************
                    //C5 data comes with special characters we must replace.
                    lin = lin.Replace("\\221", "æ");
                    lin = lin.Replace("\\233", "ø");
                    lin = lin.Replace("\\206", "å");
                    lin = lin.Replace("\\222", "Æ");
                    lin = lin.Replace("\\235", "Ø");
                    lin = lin.Replace("\\217", "Å");
                    lin = lin.Replace("\\224", "ö");
                    lin = lin.Replace("\\201", "ü");
                    lin = lin.Replace("\\204", "ä");
                    lin = lin.Replace("\\341", "ß");
                    //******************************************************
                    */
                }
                if (oldLin != null)
                    lin = oldLin + " " + lin;
                var lines = sp.Split(lin);
                if (lines.Count >= minCols)
                    return lines;
                oldLin = lin;
            }
        }

        static public double ToDouble(string str)
        {
            var val = Uniconta.Common.Utility.NumberConvert.ToDouble(str);
            if (double.IsNaN(val) || val > 9999999999999.99d || val < -9999999999999.99d) // 9.999.999.999.999,99d
                return 0d;
            return val;
        }

        //Leading zero's on GLAccount. 
        public string GLAccountFromC5(string account)
        {
            if (account != string.Empty && Accountlen != 0)
                return account.PadLeft(Accountlen, '0');
            else
                return account;
        }

        public string GLAccountFromC5_Validate(string account)
        {
            if (account == string.Empty)
                return string.Empty;

            account = this.GLAccountFromC5(account);
            if (this.LedgerAccounts.Get(account) != null)
                return account;
            return string.Empty;
        }

        public Currencies ConvertCur(string str)
        {
            if (str != string.Empty && this.CompCur != str)
            {
                Currencies cur;
                if (Enum.TryParse(str, true, out cur))
                {
                    if (cur != this.CompCurCode)
                        return cur;
                }
            }
            return 0;
        }

        public string GetVat_Validate(string vat)
        {
            if (vat != string.Empty)
            {
                var rec = (GLVat)this.vats?.Get(vat);
                if (rec != null)
                    return rec._Vat;
            }
            return string.Empty;
        }

        static public ItemUnit ConvertUnit(string str)
        {
            if (AppEnums.ItemUnit != null)
            {
                var unit = str.Replace(".", string.Empty);
                var unitId = AppEnums.ItemUnit.IndexOf(unit);
                if (unitId > 0)
                    return (ItemUnit)unitId;
            }
            return 0;
        }

        public async Task<ErrorCodes> Insert(IEnumerable<UnicontaBaseEntity> lst)
        {
            closeFile();

            if (window.Terminate)
                return ErrorCodes.NoSucces;

            var count = lst.Count();
            AppendLogLine(string.Format("Saving {0}. Number of records = {1} ...", CurFile, count));
            var lst2 = new List<UnicontaBaseEntity>();
            int n = 0;
            int nError = 0;
            ErrorCodes ret = 0;

            window.progressBar.Maximum = lst.Count();
            foreach (var rec in lst)
            {
                lst2.Add(rec);

                n++;
                if ((n % 100) == 0 || n == count)
                {
                    if (n == count)
                        window.UpdateProgressbar(count, null);
                    else
                        window.UpdateProgressbar(n * 100, null);

                    var ret2 = await api.Insert(lst2);
                    if (ret2 != Uniconta.Common.ErrorCodes.Succes)
                    {
                        foreach (var oneRec in lst2)
                        {
                            ret2 = await api.Insert(oneRec);
                            if (ret2 != 0)
                            {
                                ret = ret2;
                                nError++;
                            }
                        }
                        if (ret != 0)
                            await this.Error(ret);
                    }
                    lst2.Clear();
                }
            }
            if (nError != 0)
                AppendLogLine(string.Format("Saved {0}. Number of records with error = {1}", CurFile, nError));
            return ret;
        }
        public void Update(IEnumerable<UnicontaBaseEntity> lst)
        {
            closeFile();
            var count = lst.Count();
            AppendLogLine(string.Format("Updatating. Number of records = {0} ...", count));
            api.UpdateNoResponse(lst);

            /*
            var lst2 = new List<UnicontaBaseEntity>();
            int n = 0;
            ErrorCodes ret = 0;
            foreach (var rec in lst)
            {
                lst2.Add(rec);

                n++;
                if ((n % 100) == 0 || n == count)
                {
                    ret = await api.Update(lst2);
                    if (ret != Uniconta.Common.ErrorCodes.Succes)
                        await this.Error(ret);

                    lst2.Clear();
                }
            }
            */
        }
    }

}
