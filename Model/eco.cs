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
    public class DepartSplit
    {
        public string Afd;
        public double pct;

        static public Dictionary<string, List<DepartSplit>> importAfdelingFordeling(importLog log)
        {
            if (!log.OpenFile("AfdelingFordeling"))
            {
                return null;
            }

            try
            {
                var dim1 = log.dim1;
                List<string> lines;
                var dict = new Dictionary<string, List<DepartSplit>>();

                while ((lines = log.GetLine(3)) != null)
                {
                    var afdMaster = lines[0];
                    if (afdMaster == string.Empty)
                        continue;

                    var afd = lines[1];
                    if (dim1 != null && dim1.Get(afd) == null)
                        continue;

                    var pct = NumberConvert.ToDouble(lines[2]);
                    List<DepartSplit> lst;
                    if (dict.ContainsKey(afdMaster))
                    {
                        lst = dict[afdMaster];
                    }
                    else
                    {
                        lst = new List<DepartSplit>();
                        dict.Add(afdMaster, lst);
                    }
                    lst.Add(new DepartSplit() { Afd = afd, pct = pct });
                }
                return dict;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
                return null;
            }
        }
    }

    static public class eco
    {
        static public async Task<ErrorCodes> importAll(importLog log, int Country)
        {
            await eco.importYear(log, Country == 1 ? "RegnskabsAar" : "Regnskapsaar");
            await importDepartment(log);

            var err = await eco.importCOA(log);
            if (err != 0)
                return err;

            eco.importSystemKonti(log);

            await eco.importMoms(log, Country == 1 ? "MomsKode" : "MvaKode");
            eco.importAfgift(log, Country == 1 ? "AfgiftsKonto" : "AvgiftsKonto");

            await eco.importDebtorGroup(log);
            await eco.importCreditorGroup(log);
            await eco.importPaymentGroup(log);
            await eco.importEmployee(log);
            await eco.importDebitor(log);
            await eco.importCreditor(log);
            await eco.importInvGroup(log);
            await eco.importInv(log);

            var orders = new Dictionary<long, UnicontaBaseEntity>();
            await eco.importOrder(log, "Ordre", 1, orders);
            if (orders.Count > 0)
                await eco.importOrderLinje(log, "OrdreLinje", 1, orders);

            orders.Clear();
            await eco.importOrder(log, "Tilbud", 3, orders);
            if (orders.Count > 0)
                await eco.importOrderLinje(log, "TilbudsLinje", 3, orders);
            orders = null;

            // clear memory
            log.Items = null;
            log.ItemGroups = null;
            log.PriceLists = null;
            log.Payments = null;
            log.Employees = null;

            await eco.importContact(log);

            if (log.window.Terminate)
                return ErrorCodes.NoSucces;

            var AfdSplit = DepartSplit.importAfdelingFordeling(log);

            // we we do not import years, we do not import transaction
            if (log.years != null && log.years.Length > 0)
                await eco.importGLTrans(log, AfdSplit);

            eco.UpdateVATonAccount(log);

            return ErrorCodes.Succes;
        }

        public static async Task importDepartment(importLog log)
        {
            if (!log.OpenFile("Afdeling"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new List<GLDimType1>();
                while ((lines = log.GetLine(2)) != null)
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
                {
                    log.dim1 = new SQLCache(lst.ToArray(), true);
                    var cc = log.api.CompanyEntity;
                    cc._Dim1 = "Afdeling";
                    cc.NumberOfDimensions = 1;
                    log.api.UpdateNoResponse(cc);
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        public static async Task importEmployee(importLog log)
        {
            if (!log.OpenFile("ProjektMedarbejder"))
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
                while ((lines = log.GetLine(16)) != null)
                {
                    var rec = new Employee();
                    rec._Number = lines[0];
                    if (rec._Number == string.Empty)
                        continue;
                    rec._Title = ContactTitle.Employee;
                    rec._Name = lines[3];
                    rec._Address1 = lines[4];
                    rec._ZipCode = lines[5];
                    rec._City = lines[6];
                    if (lines[13] != string.Empty)
                        rec._Hired = DateTime.Parse(lines[13]);
                    if (lines[14] != string.Empty)
                        rec._Terminated = DateTime.Parse(lines[14]);

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.Employees = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importYear(importLog log, string Filname)
        {
            if (!log.OpenFile(Filname))
            {   
                log.AppendLogLine("Financial years not imported. No transactions will be imported");
                return;
            }
            try
            {
                List<string> lines;

                var lst = new List<CompanyFinanceYear>();
                var now = DateTime.Now;

                while ((lines = log.GetLine(4)) != null)
                {
                    var rec = new CompanyFinanceYear();
                    rec._FromDate = DateTime.Parse(lines[1]);
                    rec._ToDate = DateTime.Parse(lines[2]);
                    rec._State = FinancePeriodeState.Open; // set to open, otherwise we can't import transactions
                    rec.OpenAll();
  
                    if (rec._FromDate <= now && rec._ToDate >= now)
                        rec._Current = true;

                    log.AppendLogLine(rec._FromDate.ToShortDateString());

                    lst.Add(rec);
                }
                var err = await log.Insert(lst);
                if (err != 0)
                    log.AppendLogLine("Financial years not imported. No transactions will be imported");
                else
                    log.years = lst.ToArray();
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task<ErrorCodes> importCOA(importLog log)
        {
            if (!log.OpenFile("konto"))
            {
                return ErrorCodes.FileDoesNotExist;
            }

            try
            { 
                List<string> lines;

                var lst = new Dictionary<string, GLAccount>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(14)) != null)
                {
                    var rec = new MyGLAccount();
                    rec.SetDimUsed(1, true); // we might have afdeling
                    rec._Account = lines[0];
                    rec._Name = lines[1];
                    var t = (int)NumberConvert.ToInt(lines[2]);
                    switch (t)
                    {
                        case 5: rec._AccountType = (byte)GLAccountTypes.Header; rec._PageBreak = true; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 4: rec._AccountType = (byte)GLAccountTypes.Header; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 3: rec._AccountType = (byte)GLAccountTypes.Sum; rec._SumInfo = lines[3] + ".." + lines[0]; rec._MandatoryTax = VatOptions.NoVat; break;
                        case 1: rec._AccountType = (byte)GLAccountTypes.PL; break;
                        case 2: rec._AccountType = (byte)GLAccountTypes.BalanceSheet;
                            if (rec._Name.IndexOf("bank", StringComparison.CurrentCultureIgnoreCase) >= 0)
                            {
                                rec._MandatoryTax = VatOptions.NoVat;
                                rec._AccountType = (byte)GLAccountTypes.Bank;
                            }
                            break;
                        case 6:
                            rec._AccountType = (byte)GLAccountTypes.Sum;
                            rec._MandatoryTax = VatOptions.NoVat;
                            if (rec._Account == "6112")
                                rec._SumInfo = "1000..4990";
                            else if (rec._Account == "6199")
                                rec._SumInfo = "1000..4990;6100..6199";
                            else if (rec._Account == "8999")
                                rec._SumInfo = "1000..4990;6000..8999";
                            break;
                    }
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

        static async Task importMoms(importLog log, string Filename)
        {
            if (!log.OpenFile(Filename))
            {
                return;
            }

            try
            { 
                List<string> lines;

                var LedgerAccounts = log.LedgerAccounts;

                string lastAcc = null, lastAccOffset = null;
                var lst = new List<GLVat>();
                while ((lines = log.GetLine(5)) != null)
                {
                    var rec = new GLVat();
                    rec._Vat = lines[0];
                    rec._Name = lines[2];
                    rec._Account = lines[1];
                    rec._OffsetAccount = lines[5];
                    rec._Rate = importLog.ToDouble(lines[4]);
                    if (rec._Account != string.Empty && rec._OffsetAccount != string.Empty)
                        rec._Method = GLVatCalculationMethod.Netto;
                    else
                        rec._Method = GLVatCalculationMethod.Brutto;

                    var t = NumberConvert.ToInt(lines[3]);
                    switch (t)
                    {
                        case 1:
                        case 3:
                            rec._VatType = GLVatSaleBuy.Sales;
                            if (log.CompCountryCode == CountryCode.Norway)
                            {
                                if (rec._Rate == 25.00d)
                                    rec._TypeSales = "s3";
                                else if (rec._Rate == 15.00d)
                                    rec._TypeSales = "s4";
                                else if (rec._Rate == 8.00d)
                                    rec._TypeSales = "s5";
                                else
                                    rec._TypeSales = "s10";
                            }
                            else
                            {
                                rec._TypeSales = "s1";
                            }
                            break;
                        case 2:
                            rec._VatType = GLVatSaleBuy.Buy;
                            if (log.CompCountryCode == CountryCode.Norway)
                            {
                                if (rec._Rate == 25.00d)
                                    rec._TypeBuy = "k1";
                                else if (rec._Rate == 15.00d)
                                    rec._TypeBuy = "k2";
                                else if (rec._Rate == 8.00d)
                                    rec._TypeBuy = "k8";
                                else
                                    rec._TypeBuy = "k99";
                            }
                            else
                            {
                                rec._TypeBuy = "k1";
                            }
                            break;
                    }

                    switch(rec._Vat)
                    {
                        case "B25": rec._TypeBuy = "k3"; break;
                        case "HREP": rec._TypeBuy = "k1"; break;
                        case "REP": rec._TypeBuy = "k1"; break;
                        case "I25": rec._TypeBuy = "k1"; break;
                        case "IV25": rec._TypeBuy = "k4"; break;
                        case "IY25": rec._TypeBuy = "k5"; break;
                        case "U25": rec._TypeSales = "s1"; break;
                        case "UEUV": rec._TypeSales = "s3"; break;
                        case "UV0": rec._TypeSales = "s3"; break;
                        case "UY0": rec._TypeSales = "s4"; break;
                        case "Abr": rec._TypeSales = "s6"; break;
                    }

                    lst.Add(rec);

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
                await log.Insert(lst);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static void importAfgift(importLog log, string Filename)
        {
            if (!log.OpenFile(Filename))
            {
                return;
            }

            try
            { 
                var LedgerAccounts = log.LedgerAccounts;
                List<string> lines;

                while ((lines = log.GetLine(2)) != null)
                {
                    var rec = (MyGLAccount)LedgerAccounts.Get(lines[1]);
                    if (rec == null)
                        continue;
                    rec._SystemAccount = (byte)SystemAccountTypes.OtherTax;
                    rec.HasChanges = true;
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static void UpdateVATonAccount(importLog log)
        {
            if (log.VatIsUpdated)
                return;
            if (!log.OpenFile("konto"))
            {
                return;
            }

            try
            {
                log.AppendLogLine("Update VAT on Chart of Account");

                List<string> lines;
                var lst = new List<GLAccount>();
                var LedgerAccounts = log.LedgerAccounts;

                while ((lines = log.GetLine(14)) != null)
                {
                    var rec = (MyGLAccount)LedgerAccounts.Get(lines[0]);
                    if (rec == null)
                        continue;

                    rec._Vat = lines[4];
                    if (rec._Vat != string.Empty)
                    {
                        rec._MandatoryTax = VatOptions.Fixed;
                        rec.HasChanges = true;
                    }
                    rec._DefaultOffsetAccount = lines[6];
                    if (rec._DefaultOffsetAccount != string.Empty)
                        rec.HasChanges = true;

                    rec._PrimoAccount = lines[7];
                    if (rec._PrimoAccount != string.Empty)
                        rec.HasChanges = true;

                    if (lines[12] != "0")
                    {
                        rec.SetDimMandatory(1, true);
                        rec.HasChanges = true;
                    }
                    if (rec.HasChanges)
                        lst.Add(rec);
                }
                log.Update(lst);
                log.VatIsUpdated = true;
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static void importSystemKonti(importLog log)
        {
            if (!log.OpenFile("SystemKonto"))
            {
                return;
            }

            try
            { 
                List<string> lines;
                var lst = new List<GLAccount>();

                var LedgerAccounts = log.LedgerAccounts;

                while ((lines = log.GetLine(2)) != null)
                {
                    var rec = (GLAccount)LedgerAccounts.Get(lines[1]);
                    if (rec != null)
                    {
                        switch (lines[0])
                        {
                            case "Årsavslutning":
                            case "Årsafslutning" :
                                rec._SystemAccount = (byte)SystemAccountTypes.EndYearResultTransfer;
                                c5.ClearDimension(rec);
                                break;

                            case "Gevinst på valutakursdifferanse, kunder":
                            case "Gevinst på valutakursdifference, kunder" : rec._SystemAccount = (byte)SystemAccountTypes.ExchangeDif; break;

                            case "Feilkonto":
                            case "Fejlkonto":
                                rec._SystemAccount = (byte)SystemAccountTypes.ErrorAccount; break;
                                
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

        static async Task importDebtorGroup(importLog log)
        {
            if (!log.OpenFile("KundeGruppe"))
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
                while ((lines = log.GetLine(3)) != null)
                {
                    var rec = new DebtorGroup();
                    rec._Group = lines[0];
                    rec._Name = lines[1];
                    rec._SummeryAccount = lines[2];
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
                err = await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.DebGroups = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importCreditorGroup(importLog log)
        {
            if (!log.OpenFile("LeverandoerGruppe"))
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
                while ((lines = log.GetLine(3)) != null)
                {
                    var rec = new CreditorGroup();
                    rec._Group = lines[0];
                    rec._Name = lines[1];
                    rec._SummeryAccount = lines[2];
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
                err = await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.CreGroups = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importInvGroup(importLog log)
        {
            if (!log.OpenFile("VareGruppe"))
            {
                return;
            }

            try
            { 
                List<string> lines;

                var lst = new Dictionary<string, InvGroup>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(6)) != null)
                {
                    var rec = new InvGroup();
                    rec._Group = lines[0];
                    rec._Name = lines[1];
                    rec._RevenueAccount = lines[2];
                    rec._RevenueAccount1 = lines[3];
                    rec._RevenueAccount2 = lines[4];
                    rec._RevenueAccount3 = lines[5];
                    rec._RevenueAccount4 = rec._RevenueAccount;
                    rec._UseFirstIfBlank = true;
                    if (!lst.Any())
                        rec._Default = true;
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.ItemGroups = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importInv(importLog log)
        {
            if (!log.OpenFile("Vare"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var grpCache = log.ItemGroups;
                var lst = new Dictionary<string, InvItem>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(10)) != null)
                {
                    var rec = new InvItem();
                    rec._Item = lines[0];
                    if (string.IsNullOrEmpty(rec._Item))
                        continue;

                    rec._Name = lines[1];
                    if (grpCache.Get(lines[3]) != null)
                        rec._Group = lines[3];
                    rec._SalesPrice1 = importLog.ToDouble(lines[4]);
                    rec._CostPrice = importLog.ToDouble(lines[5]);
                    rec._Unit = importLog.ConvertUnit(lines[6]);
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.HasItem = accs.Length > 0;
                log.Items = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importOrder(importLog log, string filename, byte DCType, Dictionary<long, UnicontaBaseEntity> orders)
        {
            if (!log.OpenFile(filename))
            {
                return;
            }

            try
            {
                var Debtors = log.Debtors;

                List<string> lines;
                int MaxOrderNumber = 0;

                while ((lines = log.GetLine(40)) != null)
                {
                    var KladdeNr = NumberConvert.ToInt(lines[0]);
                    if (KladdeNr == 0)
                        continue;

                    var deb = Debtors.Get(lines[3]);
                    if (deb == null)
                        continue;

                    DCOrder rec;
                    if (DCType == 1)
                        rec = new DebtorOrder();
                    else
                        rec = new DebtorOffer();

                    rec._OrderNumber = (int)NumberConvert.ToInt(lines[1]);
                    if (rec._OrderNumber > MaxOrderNumber)
                        MaxOrderNumber = rec._OrderNumber;

                    rec._DCAccount = deb.KeyStr;
                    rec._DeliveryAddress1 = lines[11];
                    var str = string.Format("{0} {1}", lines[12], lines[13]);
                    if (str.Length > 1)
                        rec._DeliveryAddress2 = str;
                    if (lines[14] != string.Empty)
                    {
                        VatZones vatz;
                        rec._DeliveryCountry = c5.convertCountry(lines[14], log.CompCountryCode, out vatz);
                    }

                    rec._DeliveryAddress3 = lines[14];
                    if (lines[16] != string.Empty)
                        rec._DeliveryDate = DateTime.Parse(lines[16]);

                    rec._Currency = log.ConvertCur(lines[26]);

                    rec._Remark = lines[23];
                    rec._YourRef = lines[36];
                    rec._Requisition = lines[37];
                    rec._OurRef = lines[38];

                    orders.Add(KladdeNr, (UnicontaBaseEntity)rec);
                }

                await log.Insert(orders.Values);

                if (DCType == 1)
                {
                    var arr = await log.api.Query<CompanySettings>();
                    if (arr != null && arr.Length > 0)
                    {
                        arr[0]._SalesOrder = MaxOrderNumber;
                        log.api.UpdateNoResponse(arr[0]);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importOrderLinje(importLog log, string filename, byte DCType, Dictionary<long, UnicontaBaseEntity> orders)
        {
            if (!log.OpenFile(filename))
            {
                return;
            }

            try
            {
                var Items = log.Items;

                List<string> lines;

                var lst = new List<UnicontaBaseEntity>();

                while ((lines = log.GetLine(12)) != null)
                {
                    var KladdeNr = NumberConvert.ToInt(lines[0]);
                    if (KladdeNr == 0)
                        continue;

                    UnicontaBaseEntity order;
                    if (! orders.TryGetValue(KladdeNr, out order))
                        continue;

                    DCOrderLine rec;
                    if (DCType == 1)
                        rec = new DebtorOrderLine();
                    else
                        rec = new DebtorOfferLine();
                    rec._LineNumber = NumberConvert.ToInt(lines[1]);

                    if (!string.IsNullOrEmpty(lines[2]))
                    {
                        var item = (InvItem)Items.Get(lines[2]);
                        if (item == null)
                            continue;
                        rec._Item = item._Item;
                        rec._CostPrice = item._CostPrice;
                    }
                    rec._Text = lines[3];
                    rec._Qty = importLog.ToDouble(lines[4]);
                    rec._Price = importLog.ToDouble(lines[5]);
                    rec._DiscountPct = importLog.ToDouble(lines[6]);

                    rec.SetMaster(order);
                    lst.Add((UnicontaBaseEntity)rec);
                }

                await log.Insert(lst);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importPaymentGroup(importLog log)
        {
            if (!log.OpenFile("BetalingsBetingelser"))
            {
                return;
            }

            try
            {
                List<string> lines;

                var lst = new Dictionary<string, PaymentTerm>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(4)) != null)
                {
                    var rec = new PaymentTerm();
                    rec._Payment = lines[0];
                    rec._Name = lines[1];
                    rec._Days = (int)NumberConvert.ToInt(lines[3]);
                    var t = (int)NumberConvert.ToInt(lines[2]);
                    switch (t)
                    {
                        case 0: rec._PaymentMethod = PaymentMethodTypes.NetCash; rec._OffsetAccount = lines[4]; break;
                        case 1: rec._PaymentMethod = PaymentMethodTypes.NetDays; break;
                        case 2: rec._PaymentMethod = PaymentMethodTypes.EndMonth; break;
                    }
                    if (!lst.Any())
                        rec._Default = true;
                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                log.pTerms = lst.Values.ToList();
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.Payments = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importCreditor(importLog log)
        {
            if (!log.OpenFile("Leverandoer"))
            {
                return;
            }

            try
            {
                ErrorCodes err;
                List<string> lines;

                var grpCache = log.CreGroups;
                var lst = new Dictionary<string, Creditor>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(23)) != null)
                {
                    var rec = new Creditor();
                    rec._Account = lines[0];
                    if (grpCache.Get(lines[1]) != null)
                        rec._Group = lines[1];
                    rec._Name = lines[2];
                    rec._Address1 = lines[3];
                    rec._ZipCode = lines[4];
                    rec._City = lines[5];
                    rec._Address2 = lines[6];
                    rec._PaymentId = lines[12];
                    rec._LegalIdent = lines[13];
                    rec._ContactEmail = lines[16];
                    rec._PostingAccount = lines[20];
                    if (lines[22].Length != 0)
                        rec._ContactPerson = lines[22];
                    else
                        rec._ContactPerson = lines[21];
                    var t = (int)NumberConvert.ToInt(lines[8]);
                    if (t > 0)
                        rec._VatZone = (VatZones)(t - 1);

                    rec._Country = c5.convertCountry(lines[7], log.CompCountryCode, out rec._VatZone);
                    rec._Currency = log.ConvertCur(lines[9]);

                    var pay = lines[11];
                    foreach (var p in log.pTerms)
                        if (p._Name == pay)
                        {
                            rec._Payment = p._Payment;
                            break;
                        }

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);

                    /*  Now we have PostingAccount, so we do not need to update ledger
                    if (lines[20] != "" && lines[20] != lastAcc)
                    {
                        lastAcc = lines[20];
                        var acc = new GLAccount();
                        acc._Account = lastAcc;
                        err = await log.api.Read(acc);
                        if (err == 0 && acc._DefaultOffsetAccount != null)
                        {
                            acc._DefaultOffsetAccountType = GLJournalAccountType.Creditor;
                            acc._DefaultOffsetAccount = rec._Account;
                            acclst.Add(acc);
                        }
                    }
                    */
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.HasCreditor = accs.Length > 0;
                log.Creditors = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importDebitor(importLog log)
        {
            if (!log.OpenFile("Kunde"))
            {
                return;
            }

            try
            { 
                List<string> lines;

                var Employees = log.Employees;
                var grpCache = log.DebGroups;
                var lst = new Dictionary<string, Debtor>(StringNoCaseCompare.GetCompare());
                while ((lines = log.GetLine(22)) != null)
                {
                    var rec = new Debtor();
                    rec._Account = lines[0];
                    var key = grpCache.Get(lines[1]);
                    if (key != null)
                        rec._Group = key.KeyStr;
                    rec._Name = lines[2];
                    rec._Address1 = lines[3];
                    rec._ZipCode = lines[4];
                    rec._City = lines[5];
                    rec._Phone = lines[7];
                    rec._ContactEmail = lines[8];
                    rec._LegalIdent = lines[13];
                    rec._EAN = lines[14];
                    rec._ContactPerson = lines[17];
                    rec._CreditMax = importLog.ToDouble(lines[20]);
                    var t = (int)NumberConvert.ToInt(lines[11]);
                    if (t > 0)
                        rec._VatZone = (VatZones)(t - 1);

                    rec._Country = c5.convertCountry(lines[6], log.CompCountryCode, out rec._VatZone);
                    rec._Currency = log.ConvertCur(lines[12]);

                    var pay = lines[15];
                    foreach (var p in log.pTerms)
                        if (p._Name == pay)
                        {
                            rec._Payment = p._Payment;
                            break;
                        }

                    key = Employees?.Get(lines[16]);
                    if (key != null)
                        rec._Employee = key.KeyStr;

                    if (!lst.ContainsKey(rec.KeyStr))
                        lst.Add(rec.KeyStr, rec);
                }
                await log.Insert(lst.Values);
                var accs = lst.Values.ToArray();
                log.HasDebitor = accs.Length > 0;
                log.Debtors = new SQLCache(accs, true);
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importContact(importLog log)
        {
            if (!log.OpenFile("KontaktPerson"))
            {
                return;
            }

            try
            { 
                List<string> lines;

                SQLCache dcontact = log.Debtors, ccontact = log.Creditors;

                var lst = new List<Contact>();
                while ((lines = log.GetLine(17)) != null)
                {
                    var rec = new Contact();
                    rec._Name = lines[4];
                    rec._Mobil = lines[5];
                    rec._Email = lines[6];
                    if (lines[2] != "" && dcontact != null)
                    {
                        if (dcontact.Get(lines[2]) != null)
                        {
                            rec._DCAccount = lines[2];
                            rec._DCType = 1;
                        }
                    }
                    else if (lines[3] != "" && ccontact != null)
                    {
                        if (ccontact.Get(lines[3]) != null)
                        {
                            rec._DCAccount = lines[3];
                            rec._DCType = 2;
                        }
                    }
                    if (lines[10] != "0")
                        rec._Invoice = true;
                    if (lines[12] != "0")
                        rec._AccountStatement = true;
                    if (lines[13] != "0")
                        rec._InterestNote = rec._CollectionLetter = true;

                    if (string.IsNullOrWhiteSpace(rec._Name))
                        rec._Name = rec._DCAccount;
                    if (string.IsNullOrWhiteSpace(rec._Name))
                        continue;

                    lst.Add(rec);
                }
                if (lst.Count > 0)
                {
                    await log.Insert(lst);
                    log.HasContact = true;
                }
            }
            catch (Exception ex)
            {
                log.Ex(ex);
            }
        }

        static async Task importGLTrans(importLog log, Dictionary<string, List<DepartSplit>> AfdSplit)
        {
            if (!log.OpenFile("Postering"))
            {
                return;
            }

            try
            {
                List<string> lines;

                log.years = await log.api.Query<CompanyFinanceYear>();

                var department = new Dictionary<string, GLDimType1>(StringNoCaseCompare.GetCompare());
                var dim1 = log.dim1;

                var years = log.years;
                int nYears = years.Length;
                List<GLPostingLineLocal>[] YearLst = new List<GLPostingLineLocal>[nYears];
                for (int i = nYears; (--i >= 0);)
                    YearLst[i] = new List<GLPostingLineLocal>();

                var Ledger = log.LedgerAccounts;

                DateTime FirstYearStart = years[0]._FromDate;
                DateTime FirstYearEnd = years[0]._ToDate;
                int cnt = 0;
                while ((lines = log.GetLine(22)) != null)
                {
                    var posttype = lines[1].ToLower();
                    var Date = DateTime.Parse(lines[2]);
                    if (Date < FirstYearStart || (Date > FirstYearEnd && posttype.Contains("primopostering")))
                        continue;

                    bool CanSplit = false;

                    var rec = new GLPostingLineLocal();
                    rec.Date = Date;
                    rec.Account = lines[3];
                    rec.Voucher = (int)NumberConvert.ToInt(lines[4]);
                    if (rec.Voucher == 0)
                        rec.Voucher = 1;
                    rec.Text = lines[5];
                    rec.Amount = importLog.ToDouble(lines[6]);

                    var cur = log.ConvertCur(lines[7]);
                    if (cur != 0)
                    {
                        rec.Currency = cur;
                        rec.IgnoreCurDif = true; // we do not want to do any regulation in a conversion
                        rec.AmountCur = importLog.ToDouble(lines[8]);
                        if (rec.Amount * rec.AmountCur < 0d) // different sign
                        {
                            rec.Currency = null;
                            rec.AmountCur = 0;
                        }
                    }

                    if (lines[11].Length != 0) // debtor
                    {
                        rec.Settlements = "+";  // autosettle.
                        rec.Account = lines[11];
                        rec.AccountType = GLJournalAccountType.Debtor;
                        if (posttype.Contains("kundefaktura"))
                        {
                            rec.DCPostType = DCPostType.Invoice;
                            if (lines[15].Length != 0)
                                rec.DueDate = DateTime.Parse(lines[15]);
                            if (lines[13].Length != 0)
                                rec.Invoice = (int)NumberConvert.ToInt(lines[13]);
                        }
                        else
                            rec.DCPostType = DCPostType.Payment;
                    }
                    else if (lines[12].Length != 0) // creditor
                    {
                        rec.Settlements = "+";  // autosettle.
                        rec.Account = lines[12];
                        rec.AccountType = GLJournalAccountType.Creditor;
                        if (posttype.Contains("leverandørfaktura"))
                        {
                            rec.DCPostType = DCPostType.Invoice;
                            if (lines[15].Length != 0)
                                rec.DueDate = DateTime.Parse(lines[15]);
                            if (lines[14].Length != 0)
                                rec.Invoice = (int)NumberConvert.ToInt(lines[14]);
                        }
                        else
                            rec.DCPostType = DCPostType.Payment;
                    }
                    else // ledger
                    {
                        var vat = lines[16];
                        if (vat != string.Empty)
                        {
                            var acc = (MyGLAccount)Ledger.Get(rec.Account);
                            if (acc != null)
                            {
                                if (acc._MandatoryTax != VatOptions.NoVat &&
                                    acc.AccountTypeEnum != GLAccountTypes.Equity &&
                                    acc.AccountTypeEnum != GLAccountTypes.Debtor &&
                                    acc.AccountTypeEnum != GLAccountTypes.Creditor &&
                                    acc.AccountTypeEnum != GLAccountTypes.Bank &&
                                    acc.AccountTypeEnum != GLAccountTypes.LiquidAsset)
                                {
                                    acc.HasVat = true;
                                    rec.Vat = vat;
                                    rec.VatHasBeenDeducted = true;
                                }
                                CanSplit = (acc.AccountTypeEnum == GLAccountTypes.PL);
                            }
                        }
                    }

                    if (rec.AccountType > 0)
                    {
                        IdKey dc;
                        if (rec.AccountType == GLJournalAccountType.Debtor)
                            dc = log.Debtors.Get(rec.Account);
                        else
                            dc = log.Creditors.Get(rec.Account);
                        if (dc != null)
                            c5.ConvertDKText(rec, (DCAccount)dc);
                    }

                    List<GLPostingLineLocal> yearList = null;
                    for (int i = nYears; (--i >= 0);)
                    {
                        var y = years[i];
                        if (y._FromDate <= Date && y._ToDate >= Date)
                        {
                            yearList = YearLst[i];
                            yearList.Add(rec);
                            break;
                        }
                    }

                    var str = lines[17];
                    if (str != string.Empty)
                    {
                        if (AfdSplit != null && AfdSplit.ContainsKey(str))
                        {
                            if (CanSplit)
                            {
                                var amount = rec.Amount;
                                var sumAmount = 0d;
                                var amountCur = rec.AmountCur;
                                var sumAmountCur = 0d;
                                var lst = AfdSplit[str];
                                bool first = true;
                                foreach (var split in lst)
                                {
                                    if (!first)
                                    {
                                        var rec2 = new GLPostingLineLocal();
                                        rec2.Date = Date;
                                        rec2.Account = rec.Account;
                                        rec2.Voucher = rec.Voucher;
                                        rec2.Text = rec.Text;
                                        rec2.Currency = rec.Currency;
                                        rec2.Vat = rec.Vat;
                                        rec2.VatHasBeenDeducted = rec.VatHasBeenDeducted;
                                        rec = rec2;
                                        yearList.Add(rec2);
                                    }
                                    var afd = split.Afd;
                                    rec.SetDim(1, afd);
                                    if (dim1 == null && !department.ContainsKey(afd))
                                    {
                                        var dim = new GLDimType1();
                                        dim._Dim = afd;
                                        department.Add(afd, dim);
                                    }

                                    rec.Amount = Math.Round(amount * split.pct / 100d, 2);
                                    sumAmount += rec.Amount;
                                    rec.AmountCur = Math.Round(amountCur * split.pct / 100d, 2);
                                    sumAmountCur += rec.AmountCur;
                                    first = false;
                                }
                                rec.Amount += amount - sumAmount;
                                rec.AmountCur += amountCur - sumAmountCur;
                            }
                        }
                        else if (dim1 != null)
                        {
                            if (dim1.Get(str) != null)
                                rec.SetDim(1, str);
                        }
                        else
                        {
                            rec.SetDim(1, str);
                            if (!department.ContainsKey(str))
                            {
                                var dim = new GLDimType1();
                                dim._Dim = str;
                                department.Add(str, dim);
                            }
                        }
                    }
                    cnt++;
                }

                if (department.Count > 0)
                {
                    await log.Insert(department.Values);
                    log.HasDepartment = true;
                    var cc = log.api.CompanyEntity;
                    cc._Dim1 = "Afdeling";
                    cc.NumberOfDimensions = 1;
                    log.api.UpdateNoResponse(cc);
                    log.AppendLogLine(string.Format("Afdelinger {0}", department.Count));
                    department = null;
                }

                log.AppendLogLine(string.Format("Number of transactions = {0}", cnt));

                log.window.progressBar.Maximum = nYears;

                var glSort = new c5.GLTransSort();
                GLPostingHeader header = new GLPostingHeader();
                header.NumberSerie = "NR";
                header.NoDateSum = true;
                var ap = new Uniconta.API.GeneralLedger.PostingAPI(log.api);
                for (int i = 0; (i < nYears); i++)
                {
                    var arr = YearLst[i].ToArray();
                    YearLst[i] = null;
                    if (arr.Length > 0)
                    {
                        header.Comment = string.Format("Import {0} - {1} ", years[i]._FromDate.ToShortDateString(), years[i]._ToDate.ToShortDateString());

                        Array.Sort(arr, glSort);
                        c5.UpdateLedgerTrans(arr);

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

                eco.UpdateVATonAccount(log);

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
    }
}
