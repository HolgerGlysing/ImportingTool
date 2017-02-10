using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Uniconta.API.Service;
using Uniconta.API.System;
using Uniconta.Common;
using Uniconta.DataModel;
using System.IO;
using Uniconta.Common.Utility;
using Uniconta.API.GeneralLedger;
using System.Threading;
using Uniconta.ClientTools;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ImportingTool.Model
{
    static public class c5
    {
        //Get zipcode
        static public string GetZipCode(string zipcity)
        {
            if (zipcity.Contains(" "))
            {
                string zip = zipcity.Substring(0, zipcity.IndexOf(" ", StringComparison.Ordinal));
                return zip.Trim();
            }
            else
            {
                string zip = zipcity;
                if (zip.Length > 4)
                    zip = zip.Substring(0, 4);
                return zip;
            }
        }

        //Get country
        static public string GetCity(string zipcity)
        {
            if (zipcity.Contains(" "))
            {
                string city = zipcity.Substring(zipcity.IndexOf(" ", StringComparison.Ordinal) + 1);
                return city.Trim();
            }
            else
            {
                string city;
                if (zipcity.Length > 4)
                    city = zipcity.Substring(4, zipcity.Length - 4);
                else
                    city = "";
                return city;
            }
        }

        static public async Task<ErrorCodes> importAll(importLog log, int DimensionImport)
        {
            if (log.Set0InAccount)
                c5.getGLAccountLen(log);

            if (DimensionImport >= 1)
                await c5.importDepartment(log);
            if (DimensionImport >= 2)
                await c5.importCentre(log);
            if (DimensionImport >= 3)
                await c5.importPurpose(log);

            c5.importCompany(log);
            await c5.importYear(log);

            var err = await c5.importCOA(log);
            if (err != 0)
                return err;

            // we will need to update system account, since we need it in posting
            c5.importSystemKonti(log);

            err = await c5.importMoms(log);
            if (err != 0)
                return err;

            await c5.importDebtorGroup(log);
            await c5.importCreditorGroup(log);
            await c5.importEmployee(log);
            await c5.importShipment(log);
            await c5.importPaymentGroup(log);
            err = await c5.importDebitor(log);
            if (err != 0)
                return err;
            err = await c5.importCreditor(log);
            if (err != 0)
                return err;

            await c5.importContact(log, 1);
            await c5.importContact(log, 2);

            await c5.importInvGroup(log);
            await c5.importPriceGroup(log);
            var items = c5.importInv(log);
            if (items != null)
            {
                await c5.importInvPrices(log, items);
                await c5.importInvBOM(log);
            }

            var orders = new Dictionary<long, DebtorOrder>();
            await c5.importOrder(log, orders);
            if (orders.Count > 0)
                await c5.importOrderLinje(log, orders);
            orders = null;

            var purchages = new Dictionary<long, CreditorOrder>();
            await c5.importPurchage(log, purchages);
            if (purchages.Count > 0)
                await c5.importPurchageLinje(log, purchages);
            purchages = null;

            // clear memory
            log.Items = null;
            log.ItemGroups = null;
            log.PriceLists = null;
            log.Payments = null;
            log.Employees = null;

            if (log.window.Terminate)
                return ErrorCodes.NoSucces;

            // we we do not import years, we do not import transaction
            if (log.years != null && log.years.Length > 0)
            {
                DCImportTrans[] debpost = null;
                if (log.Debtors != null)
                    debpost = importDCTrans(log, true, log.Debtors);

                DCImportTrans[] krepost = null;
                if (log.Creditors != null)
                    krepost = importDCTrans(log, false, log.Creditors);

                await c5.importGLTrans(log, debpost, krepost);
            }

            // We do that after posting is called, since we have fixe vat and other things that can prevent posting
            c5.UpdateVATonAccount(log);

            return ErrorCodes.Succes;
        }

        public static async Task importYear(importLog log)
        {
            if (!log.OpenFile("exp00031", "Finansperioder"))
            {
                log.AppendLogLine("Financial years not imported. No transactions will be imported");
                return;
            }
            try
            {
                // first we load primo and ultimo periodes.
                List<List<string>> ValidLines = new List<List<string>>();

                List<string> lin;
                while ((lin = log.GetLine(8)) != null)
                {
                    if (lin[3] != "0")
                        ValidLines.Add(lin);
                }
                // then we sort it in dato order
                var ValidLines2 = ValidLines.OrderBy(rec => rec[0]);

                var lst = new List<CompanyFinanceYear>();
                DateTime fromdate = DateTime.MinValue;
                var now = DateTime.Now;

                foreach (var lines in ValidLines2)
                {
                    var datastr = lines[0];
                    int year = int.Parse(datastr.Substring(0, 4));
                    int month = int.Parse(datastr.Substring(5, 2));

                    if (lines[3] == "1")
                    {
                        fromdate = new DateTime(year, month, 1);
                    }
                    else
                    {
                        if (fromdate == DateTime.MinValue)
                            fromdate = new DateTime(year, month, 1);

                        var dayinmonth = DateTime.DaysInMonth(year, month);
                        var ToDate = new DateTime(year, month, dayinmonth);
                        if (fromdate > ToDate)
                            fromdate = ToDate.AddDays(1).AddYears(-1);

                        var rec = new CompanyFinanceYear();
                        rec._FromDate = fromdate;
                        rec._ToDate = ToDate;
                        rec._State = FinancePeriodeState.Open; // set to open, otherwise we can't import transactions
                        rec.OpenAll();

                        if (rec._FromDate <= now && rec._ToDate >= now)
                            rec._Current = true;

                        log.AppendLogLine(rec._FromDate.ToShortDateString());

                        lst.Add(rec);
                    }
                }
                var Orderedlst = lst.OrderBy(rec => rec._FromDate);
                var err = await log.Insert(Orderedlst);
                if (err != 0)
                    log.AppendLogLine("Financial years not imported. No transactions will be imported");
                else
                    log.years = Orderedlst.ToArray();
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static void getGLAccountLen(importLog log)
        {
            if (!log.OpenFile("exp00025", "Finanskonti"))
            {
                return;
            }

            try
            {
                List<string> lines;
                int AccountPrev = 0;
                int Accountlen = 0, nMax = 0;

                while ((lines = log.GetLine(2)) != null)
                {
                    var _account = log.GLAccountFromC5(lines[1]);
                    if (_account.Length >= Accountlen)
                    {
                        if (_account.Length == Accountlen)
                            nMax++;
                        else
                        {
                            int i = _account.Length;
                            while (--i >= 0)
                            {
                                var ch = _account[0];
                                if (ch < '0' || ch > '9')
                                    break;
                            }
                            if (i < 0)
                            {
                                AccountPrev = Accountlen;
                                Accountlen = _account.Length;
                                nMax = 0;
                            }
                        }
                    }
                }
                if (nMax > 2)
                    log.Accountlen = Accountlen;
                else
                    log.Accountlen = AccountPrev;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task<ErrorCodes> importCOA(importLog log)
        {
            if (!log.OpenFile("exp00025", "Finanskonti"))
            {
                return ErrorCodes.FileDoesNotExist;
            }

            try
            { 
                List<string> lines;
                var lst = new Dictionary<string, GLAccount>(StringNoCaseCompare.GetCompare());
                var sum = new List<string>();
            
                while ((lines = log.GetLine(25)) != null)
                {
                    var rec = new MyGLAccount();
                    rec._Account = log.GLAccountFromC5(lines[1]);
                    if (string.IsNullOrWhiteSpace(rec._Account))
                        continue;

                    rec._Name = lines[2];

                    var counters = lines[15].ToLower();
                    var t = GetInt(lines[3]);
                    if (t != 6 && counters != string.Empty)
                    {
                        int n = 0;
                        foreach (var r in sum)
                        {
                            if (string.Compare(r, counters) == 0)
                            {
                                rec._SaveTotal = n + 1;
                                break;
                            }
                            n++;
                        }
                        if (rec._SaveTotal == 0)
                        {
                            rec._SaveTotal = n + 1;
                            sum.Add(counters);
                        }
                    }

                    switch (t)
                    {
                        case 0: rec._AccountType = (byte)GLAccountTypes.PL; break;
                        case 1:
                            rec._AccountType = (byte)GLAccountTypes.BalanceSheet;
                            if (rec._Name.IndexOf("bank", StringComparison.CurrentCultureIgnoreCase) >= 0)
                                rec._AccountType = (byte)GLAccountTypes.Bank;
                            break;
                        case 2: rec._AccountType = (byte)GLAccountTypes.Header; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 3: rec._AccountType = (byte)GLAccountTypes.Header; rec._PageBreak = true; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 4: rec._AccountType = (byte)GLAccountTypes.Header; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 5: rec._AccountType = (byte)GLAccountTypes.Sum; rec._SumInfo = log.GLAccountFromC5(lines[10]) + ".." + rec._Account; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 6:
                            {
                                int n = 1;
                                foreach (var r in sum)
                                {
                                    counters = counters.Replace(r, string.Format("Sum({0})", n));
                                    n++;
                                }
                                rec._SumInfo = counters;
                                rec._AccountType = (byte)GLAccountTypes.CalculationExpression;
                                rec._MandatoryTax = VatOptions.NoVat;
                                break;
                            }
                    }
                    // We can't set Mandatory = true, since we can't garantie that transaction has dimensions
                    if (log.HasDepartment)
                        rec.SetDimUsed(1, true);
                    if (log.HasCentre)
                        rec.SetDimUsed(2, true);
                    if (log.HasPorpose)
                        rec.SetDimUsed(3, true);

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                var err = await log.Insert(lst.Values);
                if (err != 0)
                    return err;

                var accs = lst.Values.ToArray();
                log.LedgerAccounts = new SQLCache(accs, true);
                return ErrorCodes.Succes;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
                return ErrorCodes.Exception;
            }
        }

        static public void ClearDimension(GLAccount Acc)
        {
            Acc.SetDimUsed(1, false);
            Acc.SetDimUsed(2, false);
            Acc.SetDimUsed(3, false);
            Acc.SetDimUsed(4, false);
            Acc.SetDimUsed(5, false);
        }

        public static async Task<ErrorCodes> importMoms(importLog log)
        {
            if (!log.OpenFile("exp00014", "Momskoder"))
            {
                return ErrorCodes.Succes;
            }

            try
            {
                ErrorCodes err;
                List<string> lines;

                var LedgerAccounts = log.LedgerAccounts;

                string lastAcc = null;
                string lastAccOffset = null;
                var lst = new Dictionary<string, GLVat>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(11)) != null)
                {
                    var rec = new GLVat();
                    rec._Vat = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Vat))
                        continue;
                    rec._Name = lines[1];
                    rec._Account = log.GLAccountFromC5_Validate(lines[3]);
                    rec._OffsetAccount = log.GLAccountFromC5_Validate(lines[4]);
                    var Rate = importLog.ToDouble(lines[2]);
                    var Excempt = importLog.ToDouble(lines[5]);
                    rec._Rate = Rate * (100d - Excempt) / 100d;

                    if (rec._Account != string.Empty && rec._OffsetAccount != string.Empty)
                        rec._Method = GLVatCalculationMethod.Netto;
                    else
                        rec._Method = GLVatCalculationMethod.Brutto;

                    int vatId;
                    var vattype = lines[10].ToLower();
                    if (string.IsNullOrEmpty(vattype))
                    {
                        if (vattype.Contains("salg"))
                            vatId = 1;
                        else if (vattype.Contains("køb") || vattype.Contains("indg") || vattype.Contains("import"))
                            vatId = 2;
                        else
                            vatId = 1;
                    }
                    else
                        vatId = GetInt(vattype);

                    switch (vatId)
                    {
                        case 0:
                        case 1:
                            rec._VatType = GLVatSaleBuy.Sales;
                            rec._TypeSales = "s1";
                            break;
                        case 2:
                            rec._VatType = GLVatSaleBuy.Buy;
                            rec._TypeBuy = "k1";
                            break;
                    }

                    // Here we are assigning the VAT operations. You can see that on the VAT table on the VAT operations

                    switch (rec._Vat.ToLower())
                    {
                        case "b25": rec._TypeBuy = "k3"; rec._VatType = GLVatSaleBuy.Buy; break;
                        case "hrep":
                        case "repr":
                        case "rep": rec._TypeBuy = "k1"; rec._VatType = GLVatSaleBuy.Buy; rec._Rate = 6.25d; break;
                        case "i25": rec._TypeBuy = "k1"; rec._VatType = GLVatSaleBuy.Buy; break;
                        case "umoms":
                        case "iv25": rec._TypeBuy = "k4"; rec._VatType = GLVatSaleBuy.Buy; break;
                        case "iy25": rec._TypeBuy = "k5"; rec._VatType = GLVatSaleBuy.Buy; break;
                        case "u25": rec._TypeSales = "s1"; rec._VatType = GLVatSaleBuy.Sales; break;
                        case "ueuv": rec._TypeSales = "s3"; rec._VatType = GLVatSaleBuy.Sales; break;
                        case "uv0": rec._TypeSales = "s3"; rec._VatType = GLVatSaleBuy.Sales; break;
                        case "ueuy": rec._TypeSales = "s4"; rec._VatType = GLVatSaleBuy.Sales; break;
                        case "uy0": rec._TypeSales = "s4"; rec._VatType = GLVatSaleBuy.Sales; break;
                        case "abr": rec._TypeSales = "s6"; rec._VatType = GLVatSaleBuy.Sales; break;
                    }

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);

                    if (rec._Account != "" && rec._Account != lastAcc)
                    {
                        lastAcc = rec._Account;
                        var acc = (MyGLAccount)LedgerAccounts.Get(lastAcc);
                        if (acc != null)
                        {
                            acc._SystemAccount = rec._VatType == GLVatSaleBuy.Sales ? (byte)SystemAccountTypes.SalesTaxPayable : (byte)SystemAccountTypes.SalesTaxReceiveable;
                            acc.HasChanges = true;
                        }
                    }
                    if (rec._VatType == GLVatSaleBuy.Buy && rec._OffsetAccount != "" && rec._OffsetAccount != lastAccOffset)
                    {
                        lastAccOffset = rec._OffsetAccount;
                        var acc = (MyGLAccount)LedgerAccounts.Get(lastAccOffset);
                        if (acc != null)
                        {
                            acc._SystemAccount = (byte)SystemAccountTypes.SalesTaxOffset;
                            acc.HasChanges = true;
                        }
                    }
                }
                err = await log.Insert(lst.Values);
                if (err != 0)
                    return err;

                var vats = lst.Values.ToArray();
                log.vats = new SQLCache(vats, true);
                return 0;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
                return ErrorCodes.Exception;
            }
        }

        public static void UpdateVATonAccount(importLog log)
        {
            if (log.VatIsUpdated)
                return;

            if (!log.OpenFile("exp00025", "Finanskonti"))
            {
                return;
            }

            try
            {
                log.AppendLogLine("Update VAT on Chart of Account");

                List<string> lines;

                var LedgerAccounts = log.LedgerAccounts;
                var acclst = new List<GLAccount>();

                while ((lines = log.GetLine(16)) != null)
                {
                    var rec = (MyGLAccount)LedgerAccounts.Get(log.GLAccountFromC5(lines[1]));
                    if (rec == null)
                        continue;

                    var s = log.GetVat_Validate(lines[11]);
                    if (s != string.Empty)
                    {
                        rec._Vat = s;
                        rec.HasChanges = true;

                        // To be updated laver. We risk that we can't import transactions, since we require VAT on transactions.
                        if (lines.Count > (35 + 2) && lines[35] != "0")
                            rec._MandatoryTax = VatOptions.Fixed;
                        else
                            rec._MandatoryTax = VatOptions.Optional;
                    }
                    s = log.GLAccountFromC5_Validate(lines[8]);
                    if (s != string.Empty)
                    {
                        rec._DefaultOffsetAccount = s;
                        rec.HasChanges = true;
                    }

                    rec._Currency = log.ConvertCur(lines[13]);
                    if (rec._Currency != 0)
                        rec.HasChanges = true;

                    if (lines.Count > (36 + 2))
                    {
                        s = log.GLAccountFromC5_Validate(lines[36]);
                        if (s != string.Empty)
                        {
                            rec._PrimoAccount = s;
                            rec.HasChanges = true;
                        }
                    }
                    if (rec.HasChanges)
                        acclst.Add(rec);
                }
                log.Update(acclst);
                log.VatIsUpdated = true;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importPriceGroup(importLog log)
        {
            if (!log.OpenFile("exp00064", "Prisgrupper"))
            {
                return;
            }

            List<string> lines;

            try
            {
                var lst = new Dictionary<string, DebtorPriceList>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(3)) != null)
                {
                    var rec = new DebtorPriceList();
                    rec._Name = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Name))
                        continue;
                    rec._InclVat = lines[2] == "1";
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.PriceLists = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static void importSystemKonti(importLog log)
        {
            if (!log.OpenFile("exp00013", "SystemKonti"))
            {
                return;
            }

            try
            {
                List<string> lines;
                var lst = new List<GLAccount>();

                var LedgerAccounts = log.LedgerAccounts;

                while ((lines = log.GetLine(15)) != null)
                {
                    if (lines[1] == "0")
                    {
                        var rec = (MyGLAccount)LedgerAccounts.Get(log.GLAccountFromC5(lines[4]));
                        if (rec == null)
                            continue;

                        switch (lines[0])
                        {
                            case "ÅretsResultat" :
                            case "SYS60928":
                                rec._SystemAccount = (byte)SystemAccountTypes.EndYearResultTransfer;
                                if (log.errorAccount == null)
                                    log.errorAccount = rec.KeyStr;
                                ClearDimension(rec);
                                break;
                            case "Sek.Val.Afrunding" :
                            case "SYS60933":
                                rec._SystemAccount = (byte)SystemAccountTypes.ExchangeDif;
                                log.valutaDif = rec.KeyStr;
                                if (log.errorAccount == null)
                                    log.errorAccount = rec.KeyStr;
                                break;
                            case "Fejlkonto":
                            case "SYS11607":
                                rec._SystemAccount = (byte)SystemAccountTypes.ErrorAccount;
                                log.errorAccount = rec.KeyStr;
                                break;
                            case "Øredifference":
                            case "SYS34401":
                                rec._SystemAccount = (byte)SystemAccountTypes.PennyDif;
                                if (log.errorAccount == null)
                                    log.errorAccount = rec.KeyStr;
                                break;
                            case "AfgiftOlie":
                            case "SYS60929":
                            case "AfgiftEl":
                            case "SYS60930":
                            case "AfgiftVand":
                            case "SYS60931":
                            case "AfgiftKul":
                            case "SYS60932":
                                rec._SystemAccount = (byte)SystemAccountTypes.OtherTax;
                                rec.HasChanges = true;
                                continue;
                            default: continue;
                        }
                        lst.Add(rec);
                    }
                }
                log.Update(lst);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importDebtorGroup(importLog log)
        {
            if (!log.OpenFile("exp00034", "Debitorgrupper"))
            {
                return;
            }

            try
            {
                ErrorCodes err;
                List<string> lines;
                string lastAcc = null;
                var LedgerAccounts = log.LedgerAccounts;

                var lst = new Dictionary<string, DebtorGroup>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(10)) != null)
                {
                    var rec = new DebtorGroup();
                    rec._Group = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Group))
                        continue;
                    rec._Name = lines[1];
                    rec._SummeryAccount = log.GLAccountFromC5_Validate(lines[7]);
                    if (string.IsNullOrWhiteSpace(rec._SummeryAccount) && lastAcc != null)
                        rec._SummeryAccount = lastAcc;
                    rec._RevenueAccount = log.GLAccountFromC5_Validate(lines[2]);
                    rec._SettlementDiscountAccount = log.GLAccountFromC5_Validate(lines[8]);
                    rec._EndDiscountAccount = log.GLAccountFromC5_Validate(lines[4]);

                    rec._UseFirstIfBlank = true;
                    if (!lst.Any())
                        rec._Default = true;

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);

                    if (rec._SummeryAccount != "" && rec._SummeryAccount != lastAcc)
                    {
                        lastAcc = rec._SummeryAccount;
                        var acc = (MyGLAccount)LedgerAccounts.Get(lastAcc);
                        if (acc != null)
                        {
                            acc._AccountType = (byte)GLAccountTypes.Debtor;
                            acc.HasChanges = true;
                        }
                    }
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.DebGroups = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importCreditorGroup(importLog log)
        {
            if (!log.OpenFile("exp00042", "Kreditorgrupper"))
            {
                return;
            }

            try
            {
                ErrorCodes err;
                List<string> lines;
                string lastAcc = null;
                var LedgerAccounts = log.LedgerAccounts;

                var lst = new Dictionary<string, CreditorGroup>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(10)) != null)
                {
                    var rec = new CreditorGroup();
                    rec._Group = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Group))
                        continue;
                    rec._Name = lines[1];
                    rec._SummeryAccount = log.GLAccountFromC5_Validate(lines[7]);
                    if (string.IsNullOrWhiteSpace(rec._SummeryAccount) && lastAcc != null)
                        rec._SummeryAccount = lastAcc;
                    rec._RevenueAccount = log.GLAccountFromC5_Validate(lines[2]);
                    rec._SettlementDiscountAccount = log.GLAccountFromC5_Validate(lines[8]);
                    rec._EndDiscountAccount = log.GLAccountFromC5_Validate(lines[4]);
               
                    rec._UseFirstIfBlank = true;
                    if (!lst.Any())
                        rec._Default = true;
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);

                    if (rec._SummeryAccount != "" && rec._SummeryAccount != lastAcc)
                    {
                        lastAcc = rec._SummeryAccount;
                        var acc = (MyGLAccount)LedgerAccounts.Get(lastAcc);
                        if (acc != null)
                        {
                            acc._AccountType = (byte)GLAccountTypes.Creditor;
                            acc.HasChanges = true;
                        }
                    }
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.CreGroups = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importInvGroup(importLog log)
        {
            if (!log.OpenFile("exp00050", "Varegrupper"))
            {
                return;
            }

            try
            { 
                List<string> lines;

                var lst = new Dictionary<string, InvGroup>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(15)) != null)
                {
                    var rec = new InvGroup();
                    rec._Group = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Group))
                        continue;
                    rec._Name = lines[1];

                    rec._CostAccount = log.GLAccountFromC5_Validate(lines[3]);

                    rec._InvAccount = log.GLAccountFromC5_Validate(lines[9]);
                    rec._InvReceipt = log.GLAccountFromC5_Validate(lines[10]);
                                      
                    rec._RevenueAccount = log.GLAccountFromC5_Validate(lines[2]);
                    rec._PurchaseAccount = log.GLAccountFromC5_Validate(lines[9]);

                    rec._UseFirstIfBlank = true;

                    if (!lst.Any())
                        rec._Default = true;

                    if (! lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                var err = await log.Insert(lst.Values);
                if (err == 0)
                {
                    var accs = lst.Values.ToArray();
                    log.ItemGroups = new SQLCache(accs, true);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importShipment(importLog log)
        {
            if (!log.OpenFile("exp00019", "Forsendelser"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new Dictionary<string, ShipmentType>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(2)) != null)
                {
                    var rec = new ShipmentType();
                    rec._Number = lines[0];
                    if (rec._Number == string.Empty)
                        continue;
                    rec._Name = lines[1];

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importEmployee(importLog log)
        {
            if (!log.OpenFile("exp00011", "Medarbejder"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;

                var lst = new Dictionary<string, Employee>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(12)) != null)
                {
                    var rec = new Employee();
                    rec._Number = lines[0];
                    if (rec._Number == string.Empty)
                        continue;
                    rec._Name = lines[1];
                    rec._Address1 = lines[2];
                    rec._Address2 = lines[3];
                    rec._ZipCode = GetZipCode(lines[4]);
                    rec._City = GetCity(lines[4]);
                    rec._Mobil = lines[6];
                    rec._Email = lines[11];

                    var x = NumberConvert.ToInt(lines[10]);
                    if (x == 0 || x == 3)
                        rec._Title = ContactTitle.Employee;
                    else if (x == 2)
                        rec._Title = ContactTitle.Sales;
                    else if (x == 4)
                        rec._Title = ContactTitle.Purchase;

                    if (dim1 != null && dim1.Get(lines[8]) != null)
                        rec._Dim1 = lines[8];
                    if (dim2 != null && lines.Count > (12+2) && dim2.Get(lines[12]) != null)
                        rec._Dim2 = lines[12];
                    if (dim3 != null && lines.Count > (13 + 2) && dim3.Get(lines[13]) != null)
                        rec._Dim3 = lines[13];

                    if (lines.Count > (18 + 2) && lines[18] != string.Empty)
                        rec._Mobil = lines[18];

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                var err = await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.Employees = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static Dictionary<string, InvItem> importInv(importLog log)
        {
            if (!log.OpenFile("exp00049", "Varer"))
            {
                log.api.CompanyEntity.Inventory = false;
                return null;
            }

            try
            {
                List<string> lines;

                var grpCache = log.ItemGroups;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;
                var Creditors = log.Creditors;

                var lst = new Dictionary<string, InvItem>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(47)) != null)
                {
                    var rec = new InvItem();
                    rec._Item = lines[1];
                    if (rec._Item == string.Empty)
                        continue;

                    rec._Name = lines[2];
                    if (grpCache.Get(lines[9]) != null)
                        rec._Group = lines[9];

                    if (Creditors?.Get(lines[13]) != null)
                        rec._PurchaseAccount = lines[13];

                    if (lines[15] == "1")
                        rec._Blocked = true;

                    rec._PurchaseCurrency = (byte)log.ConvertCur(lines[7]);
                    if (rec._PurchaseCurrency != 0)
                        rec._PurchasePrice = importLog.ToDouble(lines[8]);
                    else
                        rec._CostPrice = importLog.ToDouble(lines[8]);

                    rec._Decimals = Convert.ToByte(lines[18]);
                    rec._Weight = importLog.ToDouble(lines[22]);
                    rec._Volume = importLog.ToDouble(lines[23]);
                    rec._StockPosition = lines[31];
                    rec._Unit = importLog.ConvertUnit(lines[25]);

                    //cost price unit = importLog.ToDouble(lines[43]);

                    if (dim1 != null && dim1.Get(lines[42]) != null)
                        rec._Dim1 = lines[42];
                    if (dim2 != null && lines.Count > (54 + 2) && dim2.Get(lines[54]) != null)
                        rec._Dim2 = lines[54];
                    if (dim3 != null && lines.Count > (55 + 2) && dim3.Get(lines[55]) != null)
                        rec._Dim3 = lines[55];

                    switch (lines[5])
                    {
                        case "0": rec._ItemType = (byte)ItemType.Item; break;           //Item
                        case "1": rec._ItemType = (byte)ItemType.Service; break;        //Service
                        case "2": rec._ItemType = (byte)ItemType.ProductionBOM; break;  //BOM
                        case "3": rec._ItemType = (byte)ItemType.BOM; break;            //Kit
                    }

                    switch (lines[11])
                    {
                        case "0": rec._CostModel = CostType.FIFO; break;    //FIFO
                        case "1": rec._CostModel = CostType.FIFO; break;    //LIFO
                        case "2": rec._CostModel = CostType.Fixed; break;   //Cost price
                        case "3": rec._CostModel = CostType.Average; break; //Average
                        case "4": rec._CostModel = CostType.Fixed; break;   //Serial number
                    }

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                return lst;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
            return null;
        }

        public static async Task importInvBOM(importLog log)
        {
            if (!log.OpenFile("exp00059", "Styklister"))
            {
                return;
            }

            try
            {

                var Items = log.Items;

                List<string> lines;

                var lst = new List<InvBOM>();
                while ((lines = log.GetLine(5)) != null)
                {
                    var item = Items.Get(lines[0]);
                    if (item == null)
                        continue;

                    var rec = new InvBOM();
                    rec._ItemMaster = item.KeyStr;

                    item = Items.Get(lines[2]);
                    if (item == null)
                        continue;
                    rec._ItemPart = item.KeyStr;

                    rec._LineNumber = NumberConvert.ToFloat(lines[1]);
                    rec._Qty = NumberConvert.ToFloat(lines[3]);
                    rec._QtyType = BOMQtyType.Propertional;
                    rec._UnitSize = 1;
                    rec._CostPriceFactor = 1;
                    rec._MoveType = BOMMoveType.SalesAndBuy;
  
                    lst.Add(rec);
                }
                if (lst.Count > 0)
                {
                    var err = await log.Insert(lst);
                    log.HasBOM = (err == 0);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importInvPrices(importLog log, Dictionary<string, InvItem> invItems)
        {
            List<InvPriceListLine> lst = null;
            int nPrice = 0;
            var PriceLists = log.PriceLists;
            var ItemCurFile = log.CurFile;

            if (PriceLists != null && log.OpenFile("exp00063", "Priser"))
            {
                try
                {
                    nPrice = PriceLists.Count;

                    List<string> lines;

                    InvItem item;
                    lst = new List<InvPriceListLine>();
                    while ((lines = log.GetLine(5)) != null)
                    {
                        if (!invItems.TryGetValue(lines[0], out item))
                            continue;

                        var priceList = PriceLists.Get(lines[4]);
                        if (priceList == null)
                            continue;

                        var Price = importLog.ToDouble(lines[1]);

                        var prisId = priceList.RowId;
                        if (prisId <= 3)
                        {
                            byte priceCur = (byte)log.ConvertCur(lines[3]);
                            if (prisId == 1)
                            {
                                item._SalesPrice1 = Price;
                                item._Currency1 = priceCur;
                            }
                            else if (prisId == 2)
                            {
                                item._SalesPrice2 = Price;
                                item._Currency2 = priceCur;
                            }
                            else
                            {
                                item._SalesPrice3 = Price;
                                item._Currency3 = priceCur;
                            }
                        }

                        if (nPrice > 1)
                        {
                            var rec = new InvPriceListLine();
                            rec.SetMaster((UnicontaBaseEntity)priceList);
                            rec._DCType = 1;
                            rec._Item = item._Item;
                            rec._Price = Price;
                            lst.Add(rec);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.Ex(ex);
                }
            }

            try
            {
                var PriceCurFile = log.CurFile;
                log.CurFile = ItemCurFile;

                var err = await log.Insert(invItems.Values);
                if (err == 0)
                {
                    log.HasItem = true;
                    var accs = invItems.Values.ToArray();
                    log.Items = new SQLCache(accs, true);
                }

                if (nPrice > 1)
                {
                    log.CurFile = PriceCurFile;

                    log.HasPriceList = true;
                    await log.Insert(lst);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importPaymentGroup(importLog log)
        {
            if (!log.OpenFile("exp00021", "Betalingsbetingelser"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new Dictionary<string, PaymentTerm>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(5)) != null)
                {
                    var rec = new PaymentTerm();
                    rec._Payment = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Payment))
                        continue;
                    rec._Name = lines[1];
                    rec._Days = GetInt(lines[3]);
                    var t = GetInt(lines[2]);
                    switch (t)
                    {
                        case 0: rec._PaymentMethod = PaymentMethodTypes.NetDays;  break;                    //Net
                        case 1: rec._PaymentMethod = PaymentMethodTypes.EndMonth; break;                    //End month
                        case 2: rec._PaymentMethod = PaymentMethodTypes.EndMonth; rec._Days += 90;  break;  //End quarter
                        case 3: rec._PaymentMethod = PaymentMethodTypes.EndMonth; rec._Days += 365; break;  //End year
                        case 4: rec._PaymentMethod = PaymentMethodTypes.EndWeek;  break;                    //End week
                     }
                    if (!lst.Any())
                        rec._Default = true;
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                var err = await log.Insert(lst.Values);
                if (err == 0)
                {
                    var accs = lst.Values.ToArray();
                    log.Payments = new SQLCache(accs, true);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static public CountryCode convertCountry(string str, CountryCode companyCode, out VatZones zone)
        {
            zone = VatZones.Domestic;
            if (string.IsNullOrEmpty(str))
            {
                CountryCode code;
                if (Enum.TryParse(str, true, out code))
                {
                    if (code == companyCode)
                        zone = VatZones.Domestic;
                    else
                        zone = VatZones.EU;
                    return code;
                }
                else
                {
                    switch (str)
                    {
                        case "": return companyCode;
                        case "DK":
                        case "Danmark":
                            if (companyCode != CountryCode.Denmark)
                                zone = VatZones.Foreign;
                            return CountryCode.Denmark;
                        case "S":
                        case "Sverige": zone = VatZones.EU; return CountryCode.Sweden;
                        case "N":
                        case "Norge":
                            if (companyCode != CountryCode.Norway)
                                zone = VatZones.Foreign;
                            return CountryCode.Norway;
                        case "Færøerne": zone = VatZones.EU; return CountryCode.FaroeIslands;
                        case "Grønland": zone = VatZones.Domestic; return CountryCode.Greenland;
                        case "Island": zone = VatZones.EU; return CountryCode.Iceland;
                        case "D":
                        case "Tyskland": zone = VatZones.EU; return CountryCode.Germany;
                        case "NL":
                        case "Holland": zone = VatZones.EU; return CountryCode.Netherlands;
                        case "B":
                        case "Belgien": zone = VatZones.EU; return CountryCode.Belgium;
                        case "PL":
                        case "Polen": zone = VatZones.EU; return CountryCode.Poland;
                        case "Østrig": zone = VatZones.EU; return CountryCode.Austria;
                        case "CH":
                        case "Schweiz": zone = VatZones.EU; return CountryCode.Switzerland;
                        case "F":
                        case "Frankrig": zone = VatZones.EU; return CountryCode.France;
                        case "ES":
                        case "Spanien": zone = VatZones.EU; return CountryCode.Spain;
                        case "I":
                        case "Italien": zone = VatZones.EU; return CountryCode.Italy;
                        case "Grækenland": zone = VatZones.EU; return CountryCode.Greece;
                        case "UK":
                        case "England": zone = VatZones.EU; return CountryCode.UnitedKingdom;
                        case "Rusland": zone = VatZones.Foreign; return CountryCode.Russia;
                        case "USA": zone = VatZones.Foreign; return CountryCode.UnitedStates;
                    }
                }
            }
            return companyCode;
        }

        public static async Task<ErrorCodes> importCreditor(importLog log)
        {
            if (!log.OpenFile("exp00041", "kreditorer"))
            {
                return ErrorCodes.Succes;
            }

            try
            {
                List<string> lines;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;
                var Payments = log.Payments;
                var grpCache = log.CreGroups;

                var InvoiceAccs = new List<InvoiceAccounts>();

                var lst = new Dictionary<string, Creditor>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(60)) != null)
                {
                    var rec = new Creditor();
                    rec._Account = lines[1];
                    if (string.IsNullOrWhiteSpace(rec._Account))
                        continue;

                    if (grpCache.Get(lines[11]) != null)
                        rec._Group = lines[11];
                    rec._Name = lines[2];
                    rec._Address1 = lines[3];
                    rec._Address2 = lines[4];
                    rec._ZipCode = GetZipCode(lines[5]);
                    rec._City = GetCity(lines[5]);
                    if (string.IsNullOrEmpty(lines[30]))
                    {
                        rec._PaymentMethod = PaymentTypes.VendorBankAccount;
                        rec._PaymentId = lines[30];
                    }
                    rec._LegalIdent = lines[31];
                    rec._ContactEmail = lines[57];
                    rec._ContactPerson = lines[7];
                    rec._Phone = lines[8];
                    rec._Vat = log.GetVat_Validate(lines[25]);
                    if (dim1 != null && dim1.Get(lines[32]) != null)
                        rec._Dim1 = lines[32];
                    if (dim2 != null && lines.Count > (61 + 2) && dim2.Get(lines[61]) != null)
                        rec._Dim2 = lines[61];
                    if (dim3 != null && lines.Count > (62 + 2) && dim3.Get(lines[62]) != null)
                        rec._Dim3 = lines[62];

                    rec._PricesInclVat = lines[17] != "0";

                    rec._Country = convertCountry(lines[18], log.CompCountryCode, out rec._VatZone);
                    rec._Currency = log.ConvertCur(lines[18]);

                    if (Payments?.Get(lines[20]) != null)
                        rec._Payment = lines[20];

                    if (!lst.ContainsKey(rec.KeyStr))
                    {
                        lst.Add(rec.KeyStr, rec);
                        if (lines[10] != string.Empty)
                            InvoiceAccs.Add(new InvoiceAccounts { Acc = rec._Account, InvAcc = lines[10] });
                    }
                }
                await log.Insert(lst.Values);
                log.HasCreditor = true;
                var accs = lst.Values.ToArray();
                log.Creditors = new SQLCache(accs, true);

                UpdateInvoiceAccount(log, InvoiceAccs, false);

                return ErrorCodes.Succes;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
                return ErrorCodes.Exception;
            }
        }

        public class InvoiceAccounts
        {
            public string Acc, InvAcc;
        }

        static void UpdateInvoiceAccount(importLog log, List<InvoiceAccounts> InvoiceAccs, bool isDeb)
        {
            if (InvoiceAccs.Count > 0)
            {
                var recs = new List<UnicontaBaseEntity>();
                foreach (var debUpd in InvoiceAccs)
                {
                    DCAccount rec;
                    if (isDeb)
                        rec = (DCAccount)log.Debtors.Get(debUpd.Acc);
                    else
                        rec = (DCAccount)log.Creditors.Get(debUpd.Acc);
                    if (rec != null)
                    {
                        rec._InvoiceAccount = debUpd.InvAcc;
                        recs.Add((UnicontaBaseEntity)rec);
                    }
                }
                log.Update(recs);
            }
        }

        public static async Task<ErrorCodes> importDebitor(importLog log)
        {
            if (!log.OpenFile("exp00033", "Debitorer"))
            {
                return ErrorCodes.Succes;
            }

            try
            { 
                List<string> lines;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;

                var PriceLists = log.PriceLists;
                var Employees = log.Employees;
                var Payments = log.Payments;
                var grpCache = log.DebGroups;

                var InvoiceAccs = new List<InvoiceAccounts>();

                var lst = new Dictionary<string, Debtor>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(55)) != null)
                {
                    var rec = new Debtor();
                    rec._Account = lines[1];
                    if (string.IsNullOrWhiteSpace(rec._Account))
                        continue;

                    var key = grpCache.Get(lines[11]);
                    if (key != null)
                        rec._Group = key.KeyStr;
                    rec._Name = lines[2];
                    rec._Address1 = lines[3];
                    rec._Address2 = lines[4];
                    rec._ZipCode = GetZipCode(lines[5]);
                    rec._City = GetCity(lines[5]);
                    rec._LegalIdent = lines[27];
                    rec._ContactEmail = lines[52];
                    rec._ContactPerson = lines[7];
                    rec._Phone = lines[8];
                    if (lines.Count > (58 + 2))
                        rec._EAN = lines[58];
                    rec._Vat = log.GetVat_Validate(lines[24]);
                    if (dim1 != null && dim1.Get(lines[29]) != null)
                        rec._Dim1 = lines[29];
                    if (dim2 != null && lines.Count > (56 + 2) && dim2.Get(lines[56]) != null)
                        rec._Dim2 = lines[56];
                    if (dim3 != null && lines.Count > (57 + 2) && dim3.Get(lines[57]) != null)
                        rec._Dim3 = lines[57];
                    if (lines[42] != string.Empty)
                        rec._CreditMax = importLog.ToDouble(lines[42]);

                    if (PriceLists != null)
                    {
                        var priceRec = (InvPriceList)PriceLists.Get(lines[14]);
                        if (priceRec != null)
                        {
                            rec._PriceList = priceRec.KeyStr;
                            rec._PricesInclVat = priceRec._InclVat;
                        }
                    }

                    //var t = (int)NumberConvert.ToInt(lines[11]);
                    //if (t > 0)
                    //    rec._VatZone = (VatZones)(t - 1);

                    rec._Country = convertCountry(lines[18], log.CompCountryCode, out rec._VatZone);
                    rec._Currency = log.ConvertCur(lines[18]);

                    if (Payments?.Get(lines[20]) != null)
                        rec._Payment = lines[20];

                    if (Employees?.Get(lines[23]) != null)
                        rec._Employee = lines[23];

                    if (!lst.ContainsKey(rec.KeyStr))
                    {
                        lst.Add(rec.KeyStr, rec);
                        // lines[10]  invoice account 
                        if (lines[10] != string.Empty)
                            InvoiceAccs.Add(new InvoiceAccounts { Acc = rec._Account, InvAcc = lines[10] });
                    }
                }
                await log.Insert(lst.Values);
                log.HasDebitor = true;
                var accs = lst.Values.ToArray();
                log.Debtors = new SQLCache(accs, true);

                UpdateInvoiceAccount(log, InvoiceAccs, true);

                return ErrorCodes.Succes;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
                return ErrorCodes.Exception;
            }
        }

        public static async Task importContact(importLog log, byte DCType)
        {
            string filename = (DCType == 1) ? "exp00177" : "exp00178";

            if (!log.OpenFile(filename))
            {
                log.AppendLogLine(String.Format(Localization.lookup("FileNotFound"), filename + " : KontaktPerson"));
                return;
            }

            try
            {
                SQLCache dk;
                if (DCType == 1)
                    dk = log.Debtors;
                else
                    dk = log.Creditors;

                List<string> lines;

                var lst = new List<Contact>();
                while ((lines = log.GetLine(10)) != null)
                {
                    var rec = new Contact();
                    rec._DCAccount = lines[0];
                    if (dk.Get(rec._DCAccount) != null)
                        rec._DCType = DCType;
                    else
                        rec._DCAccount = null;

                    rec._Name = lines[2];
                    if (string.IsNullOrWhiteSpace(rec._Name))
                    {
                        rec._Name = lines[0];
                        if (string.IsNullOrWhiteSpace(rec._Name))
                            continue;
                    }

                    if (lines.Count >= 13+2)
                        rec._Mobil = lines[12] != "" ? lines[12] : lines[9];
                    rec._Email = lines[8];
                   
                    if (lines[1] != "0")
                        rec._AccountStatement = rec._InterestNote = rec._CollectionLetter = rec._Invoice = true;
                    lst.Add(rec);
                }
                await log.Insert(lst);
                if (lst.Count > 0)
                    log.HasContact = true;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importOrder(importLog log, Dictionary<long, DebtorOrder> orders)
        {
            if (!log.OpenFile("exp00127", "Salgsordre"))
            {
                return;
            }

            try
            {
                var Debtors = log.Debtors;
                var Employees = log.Employees;

                List<string> lines;
                int MaxOrderNumber = 0;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;
                var Payments = log.Payments;
                var PriceLists = log.PriceLists;

                while ((lines = log.GetLine(60)) != null)
                {
                    var rec = new DebtorOrder();
                    rec._OrderNumber = (int)NumberConvert.ToInt(lines[1]);
                    if (rec._OrderNumber == 0)
                        continue;
                    if (rec._OrderNumber > MaxOrderNumber)
                        MaxOrderNumber = rec._OrderNumber;

                    var deb = (Debtor)Debtors.Get(lines[5]);
                    if (deb == null)
                        continue;
                    rec._DCAccount = deb._Account;
                    if (deb._PriceList != null)
                    {
                        var plist = (InvPriceList)PriceLists.Get(deb._PriceList);
                        rec._PricesInclVat = plist._InclVat;
                    }
                    rec._DeliveryAddress1 = lines[32];
                    rec._DeliveryAddress2 = lines[33];
                    rec._DeliveryAddress3 = lines[34];
                    if (lines[36] != string.Empty)
                    {
                        VatZones vatz;
                        rec._DeliveryCountry = c5.convertCountry(lines[36], log.CompCountryCode, out vatz);
                    }

                    if (lines[4] != string.Empty)
                        rec._DeliveryDate = GetDT(lines[4]);
                    if (lines[3] != string.Empty)
                        rec._Created = GetDT(lines[3]);

                    rec._Currency = log.ConvertCur(lines[20]);

                    rec._EndDiscountPct = importLog.ToDouble(lines[16]);
                    var val = importLog.ToDouble(lines[27]);
                    if (val <= 2)
                    {
                        rec._DeleteLines = true;
                        rec._DeleteOrder = true;
                    }

                    rec._Remark = lines[23];
                    rec._YourRef = lines[37];
                    rec._OurRef = lines[38];
                    rec._Requisition = lines[39];

                    if (Payments?.Get(lines[22]) != null)
                        rec._Payment = lines[22];

                    if (Employees?.Get(lines[25]) != null)
                        rec._Employee = lines[25];

                    if (dim1 != null && dim1.Get(lines[28]) != null)
                        rec._Dim1 = lines[28];
                    if (dim2 != null && lines.Count > (69 + 2) && dim2.Get(lines[69]) != null)
                        rec._Dim2 = lines[69];
                    if (dim3 != null && lines.Count > (70 + 2) && dim3.Get(lines[70]) != null)
                        rec._Dim3 = lines[70];

                    orders.Add(rec._OrderNumber, rec);
                }

                await log.Insert(orders.Values);

                var arr = await log.api.Query<CompanySettings>();
                if (arr != null && arr.Length > 0)
                {
                    arr[0]._SalesOrder = MaxOrderNumber;
                    log.api.UpdateNoResponse(arr[0]);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importOrderLinje(importLog log, Dictionary<long, DebtorOrder> orders)
        {
            if (!log.OpenFile("exp00128", "Ordrelinjer"))
            {
                return;
            }

            try
            {
                var Items = log.Items;
                if (Items == null)
                    return;

                List<string> lines;

                var lst = new List<DebtorOrderLine>();

                while ((lines = log.GetLine(22)) != null)
                {
                    var OrderNumber = NumberConvert.ToInt(lines[0]);
                    if (OrderNumber == 0)
                        continue;

                    DebtorOrder order;
                    if (!orders.TryGetValue(OrderNumber, out order))
                        continue;

                    var rec = new DebtorOrderLine();
                    rec._LineNumber = NumberConvert.ToInt(lines[1]);

                    InvItem item;
                    if (!string.IsNullOrEmpty(lines[2]))
                    {
                        item = (InvItem)Items.Get(lines[2]);
                        if (item == null)
                            continue;
                        rec._Item = item._Item;
                        rec._CostPrice = item._CostPrice;
                    }
                    else
                        item = null;
                    rec._Text = lines[8];
                    rec._Qty = importLog.ToDouble(lines[4]);
                    rec._Price = importLog.ToDouble(lines[5]);
                    rec._DiscountPct = importLog.ToDouble(lines[6]);
                    var amount = importLog.ToDouble(lines[7]);
                    if (rec._Price * rec._Qty == 0d && amount != 0)
                        rec._AmountEntered = amount;

                    rec._QtyNow = importLog.ToDouble(lines[11]);
                    if (rec._QtyNow == rec._Qty)
                        rec._QtyNow = 0d;

                    rec._Unit = importLog.ConvertUnit(lines[9]);
                    if (item != null && item._Unit == rec._Unit)
                        rec._Unit = 0;

                    rec._Currency = order._Currency;

                    if (rec._CostPrice == 0d)
                        rec._CostPrice = importLog.ToDouble(lines[21]);

                    rec.SetMaster(order);
                    lst.Add(rec);
                }

                await log.Insert(lst);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importPurchage(importLog log, Dictionary<long, CreditorOrder> purchages)
        {
            if (!log.OpenFile("exp00125", "Indkøbssordre"))
            {
                return;
            }

            try
            {
                var Creditors = log.Creditors;
                var Employees = log.Employees;

                List<string> lines;
                int MaxOrderNumber = 0;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;
                var Payments = log.Payments;
                var PriceLists = log.PriceLists;

                while ((lines = log.GetLine(60)) != null)
                {
                    var rec = new CreditorOrder();
                    rec._OrderNumber = (int)NumberConvert.ToInt(lines[1]);
                    if (rec._OrderNumber == 0)
                        continue;
                    if (rec._OrderNumber > MaxOrderNumber)
                        MaxOrderNumber = rec._OrderNumber;

                    var cre = Creditors.Get(lines[5]);
                    if (cre == null)
                        continue;
                    rec._DCAccount = cre.KeyStr;
                    rec._PricesInclVat = lines[16] == "1";
                    rec._DeliveryAddress1 = lines[32];
                    rec._DeliveryAddress2 = lines[33];
                    rec._DeliveryAddress3 = lines[34];
                    if (lines[4] != string.Empty)
                        rec._DeliveryDate = GetDT(lines[4]);
                    if (lines[3] != string.Empty)
                        rec._Created = GetDT(lines[3]);

                    rec._Currency = log.ConvertCur(lines[20]);

                    rec._EndDiscountPct = importLog.ToDouble(lines[17]);
                    var val = importLog.ToDouble(lines[27]);
                    if (val <= 2)
                    {
                        rec._DeleteLines = true;
                        rec._DeleteOrder = true;
                    }

                    rec._Remark = lines[23];
                    rec._YourRef = lines[37];
                    rec._OurRef = lines[38];
                    rec._Requisition = lines[39];

                    if (Payments?.Get(lines[22]) != null)
                        rec._Payment = lines[22];

                    if (Employees?.Get(lines[25]) != null)
                        rec._Employee = lines[25];

                    if (dim1 != null && dim1.Get(lines[28]) != null)
                        rec._Dim1 = lines[28];
                    if (dim2 != null && lines.Count > (69 + 2) && dim2.Get(lines[69]) != null)
                        rec._Dim2 = lines[69];
                    if (dim3 != null && lines.Count > (70 + 2) && dim3.Get(lines[70]) != null)
                        rec._Dim3 = lines[70];

                    purchages.Add(rec._OrderNumber, rec);
                }

                await log.Insert(purchages.Values);

                var arr = await log.api.Query<CompanySettings>();
                if (arr != null && arr.Length > 0)
                {
                    arr[0]._PurchaceOrder = MaxOrderNumber;
                    log.api.UpdateNoResponse(arr[0]);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importPurchageLinje(importLog log, Dictionary<long, CreditorOrder> purchages)
        {
            if (!log.OpenFile("exp00126", "Indkøbslinjer"))
            {
                return;
            }

            try
            {
                var Items = log.Items;
                if (Items == null)
                    return;

                List<string> lines;

                var lst = new List<CreditorOrderLine>();

                while ((lines = log.GetLine(22)) != null)
                {
                    var OrderNumber = NumberConvert.ToInt(lines[0]);
                    if (OrderNumber == 0)
                        continue;

                    CreditorOrder order;
                    if (!purchages.TryGetValue(OrderNumber, out order))
                        continue;

                    var rec = new CreditorOrderLine();
                    rec._LineNumber = NumberConvert.ToInt(lines[1]);

                    InvItem item;
                    if (!string.IsNullOrEmpty(lines[2]))
                    {
                        item = (InvItem)Items.Get(lines[2]);
                        if (item == null)
                            continue;
                        rec._Item = item._Item;
                    }
                    else
                        item = null;

                    rec._Text = lines[8];
                    rec._Qty = importLog.ToDouble(lines[4]);
                    rec._Price = importLog.ToDouble(lines[5]);
                    rec._DiscountPct = importLog.ToDouble(lines[6]);
                    var amount = importLog.ToDouble(lines[7]);
                    if (rec._Price * rec._Qty == 0d && amount != 0)
                        rec._AmountEntered = amount;

                    rec._QtyNow = importLog.ToDouble(lines[11]);
                    if (rec._QtyNow == rec._Qty)
                        rec._QtyNow = 0d;

                    rec._Unit = importLog.ConvertUnit(lines[9]);
                    if (item != null && item._Unit == rec._Unit)
                        rec._Unit = 0;

                    rec._Currency = order._Currency;

                    rec.SetMaster(order);
                    lst.Add(rec);
                }

                await log.Insert(lst);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static void importCompany(importLog log)
        {
            if (!log.OpenFile("exp00010", "Firmaoplysninger"))
            {
                return;
            }

            try
            {
                List<string> lines;

                if ((lines = log.GetLine(28)) != null)
                {
                    var rec = log.api.CompanyEntity;
                    rec._Address1 = lines[1];
                    rec._Address2 = lines[2];
                    rec._Address3 = lines[3];
                    rec._Phone = lines[4];
                    rec._Id = lines[21];
                    rec._NationalBank = lines[6];
                    rec._FIK = lines[7];
                    if (string.IsNullOrWhiteSpace(rec._FIK))
                        rec._FIK = lines[28];
                    rec._IBAN = lines[27];
                    rec._SWIFT = lines[25];
                    rec._FIKDebtorIdPart = 7;

                    if (log.HasDepartment)
                    {
                        rec.NumberOfDimensions = 1;
                        rec._Dim1 = "Afdeling";
                    }
                    if (log.HasCentre)
                    {
                        rec.NumberOfDimensions = 2;
                        rec._Dim2 = "Bærer";
                    }
                    if (log.HasPorpose)
                    {
                        rec.NumberOfDimensions = 3;
                        rec._Dim3 = "Formål";
                    }

                    log.NumberOfDimensions = rec.NumberOfDimensions;
                    log.api.UpdateNoResponse(rec);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importDepartment(importLog log)
        {
            if (!log.OpenFile("exp00017", "Afdelinger"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new List<GLDimType1>();
                while ((lines = log.GetLine(3)) != null)
                {
                    var rec = new GLDimType1();
                    rec._Dim = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Dim))
                        continue;
                    rec._Name = lines[1];
                    lst.Add(rec);

                    log.HasDepartment = true;
                }
                await log.Insert(lst);

                if (log.HasDepartment)
                    log.dim1 = new SQLCache(lst.ToArray(), true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importCentre(importLog log)
        {
            if (!log.OpenFile("exp00182", "Bærer"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new List<GLDimType2>();
                while ((lines = log.GetLine(2)) != null)
                {
                    var rec = new GLDimType2();
                    rec._Dim = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Dim))
                        continue;
                    rec._Name = lines[1];
                    lst.Add(rec);

                    log.HasCentre = true;
                }
                await log.Insert(lst);

                if (log.HasCentre)
                    log.dim2 = new SQLCache(lst.ToArray(), true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importPurpose(importLog log)
        {
            if (!log.OpenFile("exp00183", "Formål"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new List<GLDimType3>();
                while ((lines = log.GetLine(2)) != null)
                {
                    var rec = new GLDimType3();
                    rec._Dim = lines[0];
                    if (string.IsNullOrWhiteSpace(rec._Dim))
                        continue;
                    rec._Name = lines[1];
                    lst.Add(rec);

                    log.HasPorpose = true;
                }
                await log.Insert(lst);

                if (log.HasPorpose)
                    log.dim3 = new SQLCache(lst.ToArray(), true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static void PostDK_NotFound(importLog log, DCImportTrans[] arr, List<GLPostingLineLocal>[] YearLst, GLJournalAccountType dktype)
        {
            SQLCache Accounts = dktype == GLJournalAccountType.Debtor ? log.Debtors : log.Creditors;
            SQLCache Groups = dktype == GLJournalAccountType.Debtor ? log.DebGroups : log.CreGroups;
            var years = log.years;
            int nYears = years.Length;

            for (int n = arr.Length; (--n >= 0);)
            {
                var p = arr[n];
                if (p.Taken || p.Amount == 0)
                    continue;

                // we did not find this transaaction in the ledger. we need to generate 2 new transaction, that will net each other out
                var rec = new GLPostingLineLocal();
                rec.AccountType = dktype;
                rec.Account = p.Account;
                rec.Voucher = p.Voucher;
                rec.Text = p.Text;
                rec.Date = p.Date;
                rec.Invoice = p.Invoice;
                rec.Amount = p.Amount;
                rec.Currency = p.Currency;
                rec.AmountCur = p.AmountCur;

                var ac = (DCAccount)Accounts.Get(p.Account);
                if (ac == null)
                    continue;
                var grp = (DCGroup)Groups.Get(ac._Group);
                if (grp == null)
                    continue;

                ConvertDKText(rec, ac);

                var Offset = new GLPostingLineLocal();
                Offset.Account = grp._SummeryAccount;
                Offset.Voucher = rec.Voucher;
                Offset.Text = rec.Text;
                Offset.Date = rec.Date;
                Offset.Invoice = rec.Invoice;
                Offset.DCPostType = rec.DCPostType;
                Offset.Amount = -rec.Amount;
                Offset.Currency = rec.Currency;
                Offset.AmountCur = -rec.AmountCur;

                var Date = rec.Date;
                for (int i = nYears; (--i >= 0);)
                {
                    var y = years[i];
                    if (y._FromDate <= Date && y._ToDate >= Date)
                    {
                        var lst = YearLst[i];
                        lst.Add(rec);
                        lst.Add(Offset);

                        if (p.dif != 0d)
                        {
                            var kursDif = new GLPostingLineLocal();
                            kursDif.Date = rec.Date;
                            kursDif.Voucher = rec.Voucher;
                            kursDif.Text = rec.Text;
                            kursDif.AccountType = dktype;
                            kursDif.Account = rec.Account;
                            kursDif.Amount = p.dif;
                            kursDif.DCPostType = DCPostType.ExchangeRateDif;
                            lst.Add(kursDif);

                            kursDif = new GLPostingLineLocal();
                            kursDif.Date = rec.Date;
                            kursDif.Voucher = rec.Voucher;
                            kursDif.Text = "Ophævet kursdif fra konvertering";
                            kursDif.Account = Offset.Account;
                            kursDif.Amount = -p.dif;
                            kursDif.DCPostType = DCPostType.ExchangeRateDif;
                            lst.Add(kursDif);
                        }
                        break;
                    }
                }
            }
        }

        static public void UpdateLedgerTrans(GLPostingLineLocal[] arr)
        {
            int voucher = 0, Invoice = 0;
            DCPostType posttype = 0;
            bool clearText = false;
            string orgText = null;
            foreach (var rec in arr)
            {
                if (rec.AccountType > 0)
                {
                    Invoice = rec.Invoice;
                    voucher = rec.Voucher;
                    posttype = rec.DCPostType;
                    clearText = (rec.Text == null);
                    orgText = rec.OrgText;
                }
                else if (voucher == rec.Voucher)
                {
                    rec.Invoice = Invoice;
                    rec.DCPostType = posttype;
                    if (clearText && rec.Text == orgText)
                        rec.Text = null;
                }
            }
        }

        static public void ConvertDKText(GLPostingLineLocal rec, DCAccount dc)
        {
            var t = rec.Text;
            if (t != null && t.Length > 6)
            {
                int index = 0;

                // "Fa:13325 D:33936765" eller Kn:13328 D:97911111"
                if (t.StartsWith("Fa:"))
                    rec.DCPostType = DCPostType.Invoice;
                else if (t.IndexOf("fakt", StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    rec.DCPostType = DCPostType.Invoice;
                    index = 1;
                }
                else if (t.StartsWith("Kn:"))
                    rec.DCPostType = DCPostType.Creditnote;
                else if (t.IndexOf("kreditno", StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    rec.DCPostType = DCPostType.Creditnote;
                    index = 1;
                }
                else if (t.StartsWith("Udlign"))
                {
                    rec.DCPostType = DCPostType.Payment;
                    index = 2;
                }
                else if (t.IndexOf("betal", StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    rec.DCPostType = DCPostType.Payment;
                    index = 1;
                }
                else
                {
                    index = -1;
                    var name = ((DCAccount)dc)._Name;
                    if (name != null)
                    {
                        name = name.Replace(".", "").Replace(" ", "");
                        t = t.Replace(".", "").Replace(" ", "");
                        if (t.IndexOf(name, StringComparison.CurrentCultureIgnoreCase) >= 0 || name.IndexOf(t, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        {
                            rec.OrgText = rec.Text;
                            rec.Text = null;
                        }
                    }
                }
                if (index >= 0 && rec.DCPostType != 0)
                {
                    if (rec.Invoice == 0)
                    {
                        var tt = t.Split(' ');
                        if (tt.Length > index)
                        {
                            t = tt[index];
                            if (index == 0 && t.Length > 3)
                                t = t.Substring(3);
                            var l = t.Length;
                            if (l > 0)
                            {
                                var ch = t[l - 1];
                                if (ch == ';' || ch == '.')
                                    t = t.Substring(0, l - 1);
                                for (int lx = t.Length; (--lx >= 0);)
                                {
                                    ch = t[lx];
                                    if (ch < '0' || ch > '9')
                                    {
                                        t = null;
                                        break;
                                    }
                                }
                                if (t != null)
                                    rec.Invoice = GetInt(t);
                            }
                        }
                    }
                    if (rec.Invoice != 0)
                    {
                        rec.OrgText = rec.Text;
                        rec.Text = null;
                    }
                }
            }
        }

        public static async Task importGLTrans(importLog log, DCImportTrans[] debpost, DCImportTrans[] crepost)
        {
            if (!log.OpenFile("exp00030", "Finanspostering"))
            {
                return;
            }

            try
            {
                List<string> lines;

                log.years = await log.api.Query<CompanyFinanceYear>();

                var Ledger = log.LedgerAccounts;

                var dim1 = log.dim1;
                var dim2 = log.dim2;
                var dim3 = log.dim3;

                DCImportTrans dksearch = new DCImportTrans();
                var dkcmp = new DCImportTransSort();

                var years = log.years;
                int nYears = years.Length;
                List<GLPostingLineLocal>[] YearLst = new List<GLPostingLineLocal>[nYears];
                for (int i = nYears; (--i >= 0);)
                    YearLst[i] = new List<GLPostingLineLocal>();

                var primoPost = new List<GLPostingLineLocal>();
                DateTime primoYear = DateTime.MaxValue;

                int cnt = 0;
                while ((lines = log.GetLine(19)) != null)
                {
                    if (lines[1] != "0") //Not budget and primo.
                        continue;

                    bool IsPrimo;
                    var Date = GetDT(lines[3], out IsPrimo);
                    if (Date == DateTime.MinValue)
                        continue;

                    GLPostingLineLocal kursDif = null;

                    var rec = new GLPostingLineLocal();
                    rec.Date = Date;
                    var orgAccount = log.GLAccountFromC5(lines[0]);
                    rec.Account = orgAccount;
                    rec.Voucher = GetInt(lines[4]);
                    if (rec.Voucher == 0)
                        rec.Voucher = 1;

                    rec.Text = lines[5];
                    rec.Amount = importLog.ToDouble(lines[6]);

                    var cur = log.ConvertCur(lines[8]);
                    if (cur != 0)
                    {
                        rec.Currency = cur;
                        rec.IgnoreCurDif = true; // we do not want to do any regulation in a conversion
                        rec.AmountCur = importLog.ToDouble(lines[7]);
                        if (rec.Amount * rec.AmountCur < 0d) // different sign
                        {
                            rec.Currency = null;
                            rec.AmountCur = 0;
                        }
                    }

                    var acc = (MyGLAccount)Ledger.Get(rec.Account);
                    if (acc == null)
                        continue;

                    if (Date < primoYear) // we have an earlier date than primo year. then we mark that as the first allowed primo date.
                    {
                        primoYear = Date;
                        primoPost.Clear();
                    }

                    if (IsPrimo)
                    {
                        if (Date == primoYear) // we only take primo from first year.
                        {
                            rec.primo = 1;
                            primoPost.Add(rec);
                        }
                        continue;
                    }

                    if (acc.AccountTypeEnum == GLAccountTypes.Creditor || acc.AccountTypeEnum == GLAccountTypes.Debtor)
                    {
                        // here we convert account to debtor / creditor.
                        var arr = (acc.AccountTypeEnum == GLAccountTypes.Debtor) ? debpost : crepost;
                        if (arr != null)
                        {
                            dksearch.Amount = rec.Amount;
                            dksearch.Date = rec.Date;
                            dksearch.Voucher = rec.Voucher;

                            DCImportTrans post;
                            var idx = Array.BinarySearch(arr, dksearch, dkcmp);
                            if (idx >= 0 && idx < arr.Length)
                            {
                                post = arr[idx];
                                while (post.Taken)
                                {
                                    idx++;
                                    if (idx < arr.Length)
                                    {
                                        post = arr[idx];
                                        if (dkcmp.Compare(post, dksearch) == 0)
                                            continue;
                                    }
                                    post = null;
                                    break;
                                }
                            }
                            else
                                post = null;

                            if (post == null)
                            {
                                for (int i = arr.Length; (--i >= 0);)
                                {
                                    var p = arr[i];
                                    if (!p.Taken)
                                    {
                                        if (dkcmp.Compare(p, dksearch) == 0)
                                        {
                                            post = p;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (post != null)
                            {
                                IdKey dc;
                                if (acc.AccountTypeEnum == GLAccountTypes.Debtor)
                                {
                                    dc = log.Debtors.Get(post.Account);
                                    if (dc != null)
                                        rec.AccountType = GLJournalAccountType.Debtor;
                                }
                                else
                                {
                                    dc = log.Creditors.Get(post.Account);
                                    if (dc != null)
                                        rec.AccountType = GLJournalAccountType.Creditor;

                                }
                                if (dc != null)
                                {
                                    post.Taken = true;
                                    rec.Account = post.Account;
                                    rec.Invoice = post.Invoice;
                                    rec.Settlements = "+";  // autosettle.

                                    ConvertDKText(rec, (DCAccount)dc);

                                    if (post.dif != 0d)
                                    {
                                        kursDif = new GLPostingLineLocal();
                                        kursDif.Date = rec.Date;
                                        kursDif.Voucher = rec.Voucher;
                                        kursDif.Text = rec.Text;
                                        kursDif.AccountType = rec.AccountType;
                                        kursDif.Account = rec.Account;
                                        kursDif.Amount = post.dif;
                                        kursDif.DCPostType = DCPostType.ExchangeRateDif;

                                        post.dif = 0d;
                                    }
                                }
                            }
                        }
                    }
                    else if (acc._MandatoryTax != VatOptions.NoVat && 
                        acc.AccountTypeEnum != GLAccountTypes.Equity && 
                        acc.AccountTypeEnum != GLAccountTypes.Bank && 
                        acc.AccountTypeEnum != GLAccountTypes.LiquidAsset)
                    {
                        rec.Vat = log.GetVat_Validate(lines[9]);
                        if (rec.Vat != string.Empty)
                            acc.HasVat = true;
                        rec.VatHasBeenDeducted = true;
                    }

                    if (dim1 != null && lines[2].Length != 0)
                    {
                        var str = lines[2];
                        if (dim1.Get(str) != null)
                            rec.SetDim(1, str);
                    }
                    if (dim2 != null && lines.Count > (19+2) && lines[19].Length != 0)
                    {
                        var str = lines[19];
                        if (dim2.Get(str) != null)
                            rec.SetDim(2, str);
                    }
                    if (dim3 != null && lines.Count > (20 + 2) && lines[20].Length != 0)
                    {
                        var str = lines[20];
                        if (dim3.Get(str) != null)
                            rec.SetDim(3, str);
                    }

                    for (int i = nYears; (--i >= 0);)
                    {
                        var y = years[i];
                        if (y._FromDate <= Date && y._ToDate >= Date)
                        {
                            YearLst[i].Add(rec);
                            if (kursDif != null)
                            {
                                YearLst[i].Add(kursDif);
                                rec = new GLPostingLineLocal();
                                rec.Date = kursDif.Date;
                                rec.Voucher = kursDif.Voucher;
                                rec.Text = "Ophævet kursdif fra konvertering";
                                rec.Account = orgAccount;
                                rec.Amount = -kursDif.Amount;
                                rec.DCPostType = DCPostType.ExchangeRateDif;
                                YearLst[i].Add(rec);
                            }
                            break;
                        }
                    }
                    cnt++;
                }

                if (primoPost.Any())
                {
                    for (int i = nYears; (--i >= 0);)
                    {
                        var y = years[i];
                        if (y._FromDate <= primoYear && y._ToDate >= primoYear)
                        {
                            YearLst[i].AddRange(primoPost);
                            cnt += primoPost.Count;
                            primoPost = null;
                            break;
                        }
                    }
                }

                log.AppendLogLine(string.Format("Number of transactions = {0}", cnt));

                if (debpost != null)
                    PostDK_NotFound(log, debpost, YearLst, GLJournalAccountType.Debtor);
                if (crepost != null)
                    PostDK_NotFound(log, crepost, YearLst, GLJournalAccountType.Creditor);

                log.window.progressBar.Maximum = nYears;

                var glSort = new GLTransSort();
                GLPostingHeader header = new GLPostingHeader();
                header.NumberSerie = "NR";
                header.NoDateSum = true; // 
                var ap = new Uniconta.API.GeneralLedger.PostingAPI(log.api);
                for (int i = 0; (i < nYears); i++)
                {
                    if (log.errorAccount != null)
                    {
                        long sum = 0;
                        foreach (var rec in YearLst[i])
                            sum += NumberConvert.ToLong(rec.Amount * 100d);

                        if (sum != 0)
                        {
                            var rec = new GLPostingLineLocal();
                            rec.Date = years[i]._ToDate;
                            rec.Account = log.GLAccountFromC5(log.errorAccount);
                            rec.Voucher = 99999;
                            rec.Text = "Ubalance ved import fra C5";
                            rec.Amount = sum / -100d;
                            YearLst[i].Add(rec);
                        }
                    }
                    var arr = YearLst[i].ToArray();
                    YearLst[i] = null;
                    if (arr.Length > 0)
                    {
                        header.Comment = string.Format("Import {0} - {1}", years[i]._FromDate.ToShortDateString(), years[i]._ToDate.ToShortDateString());

                        Array.Sort(arr, glSort);
                        UpdateLedgerTrans(arr);

                        var res = await ap.PostJournal(header, arr, false);
                        if (res.Err != 0)
                        {
                            log.AppendLogLine("Fejl i poster i " + header.Comment);
                            await log.Error(res.Err);
                        }
                        else
                            log.AppendLogLine("Posting " + header.Comment);
                    }
                    log.window.UpdateProgressbar(i, null);
                }

                foreach (var ac in (MyGLAccount[])Ledger.GetNotNullArray)
                {
                    if (!ac.HasVat && ac._MandatoryTax != VatOptions.NoVat)
                    {
                        ac._MandatoryTax = VatOptions.NoVat;
                        ac.HasChanges = true;
                    }
                }

                c5.UpdateVATonAccount(log);

                log.AppendLogLine("Generate Primo Transactions");

                var ap2 = new FinancialYearAPI(ap);
                for (int i = 1; (i < nYears); i++)
                {
                    await ap2.GeneratePrimoTransactions(years[i], null, null, 9999, "NR");
                }
                log.window.UpdateProgressbar(0, "Done");

            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public class GLTransSort : IComparer<GLPostingLineLocal>
        {
            public int Compare(GLPostingLineLocal x, GLPostingLineLocal y)
            {
                int c = DateTime.Compare(x.Date, y.Date);
                if (c != 0)
                    return c;
                c = y.primo - x.primo;
                if (c != 0)
                    return c;
                c = x.Voucher - y.Voucher;
                if (c != 0)
                    return c;
                c = (int)y.AccountType - (int)x.AccountType;
                if (c != 0)
                    return c;
                var v = x.Amount - y.Amount;
                if (v > 0.001d)
                    return 1;
                if (v < -0.001d)
                    return 1;
                return 0;
            }
        }

        public class DCImportTransSort : IComparer<DCImportTrans>
        {
            public int Compare(DCImportTrans x, DCImportTrans y)
            {
                if (x.Date > y.Date)
                    return 1;
                if (x.Date < y.Date)
                    return -1;
                int c = x.Voucher - y.Voucher;
                if (c != 0)
                    return c;
                var v = x.Amount - y.Amount;
                if (v > 0.001d)
                    return 1;
                if (v < -0.001d)
                    return 1;
                return 0;
            }
        }

        public class DCImportTrans
        {
            public DateTime Date;
            public DateTime DueDate;
            public DateTime DocumentDate;
            public double Amount;
            public double AmountCur;
            public double dif;
            public Currencies? Currency;
            public string Account;
            public string Text;
            public int Voucher;
            public int Invoice;
            public bool Taken;
        }

        public static int GetInt(string str)
        {
            var l = str.Length;
            if (l == 0)
                return 0;
            if (l > 8)
                str = str.Substring(l - 8);
            return (int)NumberConvert.ToInt(str);
        }

        public static DateTime GetDT(string str)
        {
            bool IsPrimo;
            return GetDT(str, out IsPrimo);
        }

        public static DateTime GetDT(string str, out bool IsPrimo)
        {
            IsPrimo = false;
            if (str.Length < 10)
                return DateTime.MinValue;

            int year = int.Parse(str.Substring(0, 4));
            int month = int.Parse(str.Substring(5, 2));
            var daystr = str.Substring(8, 2);
            int dayinmonth = 0;
            if (daystr == "PR")
            {
                IsPrimo = true;
                dayinmonth = 1;
            }
            else if (daystr == "UL")
                dayinmonth = DateTime.DaysInMonth(year, month);
            else
                dayinmonth = int.Parse(daystr);
            return new DateTime(year, month, dayinmonth);
        }

        public static DCImportTrans[] importDCTrans(importLog log, bool deb, SQLCache Accounts)
        {
            if (deb)
            {
                if (!log.OpenFile("exp00037", "Debitorposteringer"))
                    return null;
            }
            else
            {
                if (!log.OpenFile("exp00045", "Kreditorposteringer"))
                    return null;
            }

            try
            {
                List<string> lines;

                int offset = deb ? 0 : 1;  // kreditor has one less
                List<DCImportTrans> lst = new List<DCImportTrans>();
                while ((lines = log.GetLine(29)) != null)
                {
                    var Date = GetDT(lines[3]);
                    if (Date == DateTime.MinValue)
                        continue;

                    var ac = Accounts.Get(lines[1]);
                    if (ac == null)
                        continue;

                    var rec = new DCImportTrans();
                    rec.Date = Date;
                    rec.Account = ac.KeyStr;
                    rec.Invoice = GetInt(lines[deb ? 4 : 22]);
                    rec.Voucher = GetInt(lines[5 - offset]);
                    if (rec.Voucher == 0)
                        rec.Voucher = 1;
                    rec.Text = lines[6 - offset];
                    rec.Amount = importLog.ToDouble(lines[8 - offset]);
                    rec.dif = importLog.ToDouble(lines[22 - offset]);

                    var cur = log.ConvertCur(lines[10 - offset]);
                    if (cur != 0)
                    {
                        rec.Currency = cur;
                        rec.AmountCur = importLog.ToDouble(lines[9 - offset]);
                        if (rec.Amount * rec.AmountCur < 0d) // different sign
                        {
                            rec.Currency = null;
                            rec.AmountCur = 0;
                        }
                    }
                    lst.Add(rec);
                }
                if (lst.Count() > 0)
                {
                    var arr = lst.ToArray();
                    Array.Sort(arr, new DCImportTransSort());
                    return arr;
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
            return null;
        }
    }
}
