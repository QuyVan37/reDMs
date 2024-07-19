using Antlr.Runtime.Misc;
using Aspose.Slides;
using Aspose.Slides.Export.Web;
using Avalonia.Controls;
using DMS03.Models;
using Microsoft.Ajax.Utilities;
using Microsoft.Office.Core;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SqlTypes;
using System.Deployment.Internal;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.RightsManagement;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls.WebParts;
using static iTextSharp.text.pdf.AcroFields;

namespace DMS03.Controllers
{
    public class CommonFunctionController : Controller
    {
        private DMSEntities db = new DMSEntities();
        private CVNHolidays dbHolidays = new CVNHolidays();

        public SendEmailController sendEmailJob = new SendEmailController();
        public StoreProcudure storeProcudure = new StoreProcudure();
        // GET: CommonFunction
        public bool isNeedPLAN(int MRtype, double price, string unit)
        {
            bool result = true;
            if (MRtype == 6) // X7-repair
            {
                result = false;
            }
            else
            {
                // chuyển đổi tiền tệ sang VND.
                double VNDPrice = 0;
                double RateVNDtoUSD = db.ExchangeRates.ToList().LastOrDefault().RateVNDtoUSD.Value;
                double RateJPYtoUSD = db.ExchangeRates.ToList().LastOrDefault().RateJPYtoUSD.Value;
                if (unit == "USD")
                {
                    VNDPrice = price * RateVNDtoUSD;
                }
                if (unit == "JPY")
                {
                    var UsdPrice = price / RateJPYtoUSD;
                    VNDPrice = UsdPrice * RateVNDtoUSD;
                }
                if (unit == "VND")
                {
                    VNDPrice = price;
                }
                if (VNDPrice >= 30000000) //30 triêu VND
                {
                    result = true;
                }
                else
                {
                    result = false;
                }

            }

            return result;
        }
        public string[] AutoAccFill(int MRtype, string ModelColorOrMono, string supplierCode, double price, string unit, string firt2LetterOfPartNo)
        {

            string G_L = "";
            string location = "";
            string assetNumber = "";
            bool isBigger30tr = isNeedPLAN(MRtype, price, unit);
            // 1.GL Account:
            //Nếu X7  => A214-307800-3120-3200
            if (MRtype == 6) //X7-Repair
            {
                G_L = "A214-307800-3120";
            }
            else
            {
                if (isBigger30tr == true) // Nếu khác X7 và price >30trVND => 2170-001000
                {
                    G_L = "2170-001000";
                }
                else// Other(else) => A212-103000-3120-3200
                {
                    G_L = "A212-103000-3120";
                }
                if (price == 10)
                {
                    G_L = "-";
                }
            }

            // 2. Location
            // nếu Inhouse (Code: 5400 || 5500)
            if (!String.IsNullOrEmpty(ModelColorOrMono))
            {
                if (supplierCode == "5400" || supplierCode == "5500" || supplierCode == "3400" || supplierCode == "3500")
                {
                    if (ModelColorOrMono.ToUpper().Contains("MONO"))
                    {
                        location = "964-MONO";
                    }
                    else
                    {
                        location = "965-COLOR";
                    }
                }
                else
                {
                    if (ModelColorOrMono.ToUpper().Contains("MONO"))
                    {
                        location = "961-MONO";
                    }
                    else
                    {
                        location = "963-COLOR";
                    }
                }
            }
            else
            {
                location = "-";
            }


            // 3.AssetNumber
            if (firt2LetterOfPartNo.Contains("RX") || firt2LetterOfPartNo.Contains("FX") || firt2LetterOfPartNo.Contains("RJ")) // Packing part
            {
                assetNumber = "08140P";
            }
            else
            {
                assetNumber = "08110P";
            }


            string[] result = new string[] { G_L, location, assetNumber };
            return result;
        }

        public void autoUpdateAccFill()
        {
            if (Session["Admin"].ToString() == "Admin")
            {
                Task.Factory.StartNew(() =>
                {
                    var listMRNotInputLocation = db.MRs.Where(x => x.Location == "-" && x.Belong != "CRG").ToList();
                    foreach (var item in listMRNotInputLocation)
                    {

                        var modelColorOrMono = db.ModelLists.Find(item.ModelID).ModelType;
                        var supplierCode = db.Suppliers.Find(item.SupplierID).SupplierCode;
                        var first2LeterPartNo = item.PartNo.Remove(2, item.PartNo.Length - 2);
                        string[] accInfor = AutoAccFill(item.TypeID.Value, modelColorOrMono, supplierCode, item.EstimateCost.Value, item.Unit, first2LeterPartNo);
                        item.GLAccount = accInfor[0];
                        item.Location = accInfor[1];
                        item.AssetNumber = accInfor[2];

                        if (item.Location == "-")
                        {
                            string msg = "You not yet setup Model Type : Mono/Color for model " + db.ModelLists.Find(item.ModelID).ModelName;
                            sendEmailJob.sendEmailToAdmin(msg);
                        }

                        db.Entry(item).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                });
            }

        }


        public string isRenewOrAddOrMT(string DieNo)
        {
            var result = ""; // >> MT || Renewal || Additional || Overhaul || Invalid

            var dieNo = DieNo == null ? "" : DieNo.Trim();


            if (dieNo.Length < 3) return "Invalid";
            var fistLetter = dieNo[0];
            var seccondLetter = dieNo[1];
            if (dieNo.Length == 3)
            {
                if (seccondLetter == '4' || seccondLetter == '3' || seccondLetter == '2')
                {
                    result = "Renewal";
                }
                else
                {

                    if (seccondLetter == '1')
                    {
                        if (fistLetter == '1')
                        {
                            result = "MT";
                        }
                        else
                        {
                            result = "Additional";
                        }
                    }
                    else
                    {
                        result = "Invalid";
                    }

                }
            }
            else
            {
                result = "Invalid";
            }
            return result;
        }
        public int? getNummberInString(string str_num)
        {

            string numbersOnly = Regex.Replace(str_num, "[^0-9]", "");
            int num = 0;
            bool isNum = int.TryParse(numbersOnly, out num);
            int? result = isNum ? num : new Nullable<Int32>();
            return result;

        }
        public bool isCanSeePrice(int userID)
        {
            return true;
        }
        public bool checkMrExistOrNot(string partNo, string clasification, string drawHis, int? mRType, int mRID)
        {
            bool result = false;
            // true :da ton tai => ko duoc phep them vao
            // false : chua ton tai => dc phep them vao
            MR checkExist = null;
            if (mRType == 4) // X5 
            {
                //Trương hợp X5 phải check cả his
                // Cùng 1 his ko dược trùng Dim
                checkExist = db.MRs.AsNoTracking().Where(x => x.PartNo == partNo.Trim().ToUpper() && x.Clasification == clasification && x.DrawHis == drawHis && x.MRID != mRID && x.Active != false && (x.StatusID != 11 && x.StatusID != 12)).FirstOrDefault();
            }
            else
            {
                checkExist = db.MRs.AsNoTracking().Where(x => x.PartNo == partNo.Trim().ToUpper() && x.Clasification == clasification && x.MRID != mRID && x.Active != false && (x.StatusID != 11 && x.StatusID != 12)).FirstOrDefault();
            }
            if (checkExist != null)
            {
                result = true;

            }
            return result;
        }

        public class checkDie
        {
            public bool isExist { set; get; }
            public bool isCancel { set; get; }
            public bool isClose { set; get; }
            public bool isOfficial { set; get; }
            public int NoOfRecord { set; get; }
            public int DieID { set; get; }
        }
        public checkDie checkDieExist(string dieNo)
        {
            var die = db.Die1.Where(x => x.DieNo == dieNo.ToUpper().Trim() && x.Active != false);
            checkDie output = new checkDie();
            if (die.Count() > 0)
            {
                output.isExist = true;
                output.isCancel = die.FirstOrDefault().isCancel == true ? true : false;
                output.isClose = die.FirstOrDefault().isClosed == true ? true : false;
                output.isOfficial = die.FirstOrDefault().isOfficial == true ? true : false;
                output.NoOfRecord = die.Count();
                output.DieID = die.FirstOrDefault().DieID;
            }
            else
            {
                output.isExist = false;
                output.isCancel = false;
                output.isClose = false;
                output.isOfficial = false;
                output.DieID = 0;
            }

            return output;
        }

        public int genarateNewDie(string dieNo, string dieCode, string partNo, string processCodeID, string ModelID, string SupplierID, string needUseDate, string targetOKDate, string dept, MR mR, int? warranty)
        {
            Die1 newDie = new Die1();
            int NEWDIEID = 0;
            int NEWPARTID = 0;


            // Create from MR
            if (mR != null)
            {

                string[] cfig = { "MT", "Renewal", "Additional" };
                if (cfig.Contains(isRenewOrAddOrMT(mR.Clasification))) // X1 new, X1-add, X4-Add
                {
                    // add Die
                    checkDie checkDieResult = checkDieExist(mR.DieNo);
                    if (checkDieResult.isExist)
                    {
                        newDie = db.Die1.Find(checkDieResult.DieID);
                    }
                    newDie.DieNo = mR.DieNo;
                    newDie.Die_Code = mR.Clasification;
                    newDie.SpecialSpec = newDie.SpecialSpec + mR.DieSpecial;
                    newDie.PartNoOriginal = mR.PartNo;
                    newDie.ProcessCodeID = mR.ProcessCodeID.Value;
                    newDie.ModelID = mR.ModelID.Value;
                    newDie.DieMaker = db.Suppliers.Where(x => x.SupplierCode == mR.OrderTo).FirstOrDefault().SupplierName;
                    newDie.SupplierID = mR.SupplierID.Value;
                    newDie.MCsize = mR.MCSize;
                    newDie.CavQuantity = mR.CavQty;
                    newDie.DieCost_USD = mR.AppCostExchangeUSD;
                    newDie.DieWarranty_Short = warranty;
                    newDie.PODate = mR.PODate;
                    newDie.DieClassify = isRenewOrAddOrMT(mR.Clasification);
                    newDie.DieMakeLocation = mR.MakeLocation;
                   // newDie.DieMaker = mR.DieMaker;
                    newDie.DieStatusID = 1; // Under Making
                    newDie.MR_Request_Date = mR.RequestDate;
                    newDie.MR_App_Date = mR.PURAppDate;
                    newDie.PO_Issue_Date = db.PO_Dies.Where(x => x.MRID == mR.MRID && x.Active != false && x.POStatusID != 20).FirstOrDefault()?.IssueDate;
                    newDie.PODate = mR.PODate;
                    newDie.Active = true;
                    newDie.isCancel = false;
                    newDie.isClosed = false;
                    newDie.isOfficial = true;
                    newDie.Belong = mR.Belong;
                    // newDie = updateDieStatus(newDie);
                    if (checkDieResult.isExist)
                    {
                        db.Entry(newDie).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    else
                    {
                        newDie.IssueDate = DateTime.Now;
                        newDie.Decision_Date = DateTime.Now.ToShortDateString();
                        db.Die1.Add(newDie);
                        db.SaveChanges();
                    }
                    NEWDIEID = newDie.DieID;
                    //***** => DIEID

                    // Add Part
                    NEWPARTID = addNewPart(mR.PartNo, mR.PartName, mR.ModelID.ToString(), mR.CommonPart + "," + mR.FamilyPart);
                    //**** => PARTID


                    // Add CommonDie1
                    addCommonDie(NEWDIEID, mR.DieNo, NEWPARTID, mR.PartNo);


                    // Nếu có nhiều die thành phần thì phải tạo ra nhiều record và tăng đuôi
                    if (mR.NoOfDieComponent > 1)
                    {
                        //mR.DieNo; ==>  RC5-1234-000-11A-001
                        var mainDieNo = mR.DieNo.Remove(mR.DieNo.Length - 3); // RC5-1234-000-11A-
                        for (var i = 2; i <= mR.NoOfDieComponent; i++)
                        {
                            var tail = i;
                            var tailStr = "";
                            if (tail.ToString().Length == 1)
                            {
                                tailStr = "00" + tail.ToString();
                            }
                            if (tail.ToString().Length == 2)
                            {
                                tailStr = "0" + tail.ToString();
                            }

                            Die1 newDie1 = new Die1();
                            db.Entry(newDie).State = EntityState.Detached;
                            newDie1 = newDie;
                            newDie1.DieID = 0;
                            newDie1.DieNo = mainDieNo + tailStr;

                            // add Die
                            checkDie checkDieResult1 = checkDieExist(newDie1.DieNo);
                            if (checkDieResult1.isExist)
                            {
                                newDie1 = db.Die1.Find(checkDieResult1.DieID);
                            }
                            else
                            {
                                db.Entry(newDie1).State = EntityState.Added;
                                db.SaveChanges();
                                addCommonDie(newDie1.DieID, newDie1.DieNo, NEWPARTID, mR.PartNo);
                            }
                        }
                    }
                    //New Part has Family common
                    if (mR.CommonPart != null)
                    {
                        if (mR.CommonPart.Trim().Length > 7)
                        {
                            string[] arrListPart = mR.CommonPart.Split(',');
                            foreach (var item in arrListPart)
                            {
                                int newCommonPartID = addNewPart(item.Trim().ToUpper(), null, mR.ModelID.ToString(), mR.CommonPart + "," + mR.FamilyPart);
                                addCommonDie(NEWDIEID, mR.DieNo, newCommonPartID, item.Trim().ToUpper());
                            }
                        }
                    }

                    if (mR.FamilyPart != null)
                    {
                        if (mR.FamilyPart.Trim().Length > 7)
                        {
                            string[] arrListPart = mR.FamilyPart.Split(',');
                            foreach (var item in arrListPart)
                            {
                                int newCommonPartID = addNewPart(item.Trim().ToUpper(), null, mR.ModelID.ToString(), mR.CommonPart + "," + mR.FamilyPart);
                                addCommonDie(NEWDIEID, mR.DieNo, newCommonPartID, item.Trim().ToUpper());
                            }
                        }
                    }
                }
            }
            else // Create Die from addNewItem
            {
                // add Die
                checkDie checkDieResult = checkDieExist(dieNo);
                if (checkDieResult.isExist)
                {
                    newDie = db.Die1.Find(checkDieResult.DieID);
                }
                newDie.DieNo = dieNo;
                newDie.Die_Code = dieCode;
                newDie.PartNoOriginal = partNo;
                newDie.ProcessCodeID = !String.IsNullOrWhiteSpace(processCodeID) ? int.Parse(processCodeID) : newDie.ProcessCodeID;
                newDie.ModelID = !String.IsNullOrWhiteSpace(ModelID) ? int.Parse(ModelID) : newDie.ModelID;
                newDie.SupplierID = !String.IsNullOrWhiteSpace(SupplierID) ? int.Parse(SupplierID) : newDie.SupplierID;
                newDie.DieClassify = isRenewOrAddOrMT(dieCode);
                DateTime cvDate = new DateTime();
                newDie.Need_Use_Date = DateTime.TryParse(needUseDate, out cvDate) ? needUseDate : newDie.Need_Use_Date;
                newDie.Target_OK_Date = DateTime.TryParse(targetOKDate, out cvDate) ? cvDate : newDie.Target_OK_Date;
                newDie.DieStatusID = 1; // Under Making
                newDie.Active = true;
                newDie.isCancel = false;
                newDie.isClosed = false;
                newDie.isOfficial = false;
                newDie.Belong = dept.ToString().Contains("CRG") ? "CRG" : "LBP";
                // newDie = updateDieStatus(newDie);
                if (checkDieResult.isExist)
                {
                    db.Entry(newDie).State = EntityState.Modified;
                    db.SaveChanges();
                }
                else
                {
                    newDie.IssueDate = DateTime.Now;
                    newDie.Decision_Date = DateTime.Now.ToShortDateString();
                    db.Die1.Add(newDie);
                    db.SaveChanges();
                }
                NEWDIEID = newDie.DieID;
                //***** => DIEID 

                // New Part
                int newCommonPartID = addNewPart(partNo.Trim().ToUpper(), null, ModelID, null);
                addCommonDie(NEWDIEID, dieNo, newCommonPartID, partNo);
                //**** => PARTID
            }

            return NEWDIEID;
        }

        public void cancelDie(MR mR)
        {
            var today = DateTime.Now;
            if (mR.TypeID == 1 || mR.TypeID == 2 || mR.TypeID == 3) // MR New/Add/Renew
            {
                var existDieLaunching = db.Die1.Where(x => x.DieNo == mR.DieNo && x.Active != false).FirstOrDefault();
                if (existDieLaunching != null)
                {
                    existDieLaunching.isCancel = true;
                    existDieLaunching.isOfficial = false;
                    existDieLaunching.Genaral_Information = today.ToString("yyyy-MM-dd") + "MR was cancel this die " + System.Environment.NewLine + existDieLaunching.Genaral_Information;
                    db.Entry(existDieLaunching).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
        }
        public void cancelTPI(MR mR, bool? cancel_NoPay, bool? Cancel_NoRepair, string name, string EmailSender)
        {
            var today = DateTime.Now;
            if (!String.IsNullOrWhiteSpace(mR.TroubleID))
            {
                var arrTroubleID = mR.TroubleID?.Split(',');
                foreach (var id in arrTroubleID)
                {
                    var trb = db.Troubles.Find(int.Parse(id));
                    trb.CloseContent = today.ToString("yyyy-MM-dd ") + name + " Was cancel MR";
                    trb.CloseDate = today;
                    if (cancel_NoPay == true)
                    {
                        trb.FinalStatusID = 6; // W-FA
                    }
                    if (Cancel_NoRepair == true)
                    {
                        trb.FinalStatusID = 15; // Cancel
                        trb.CloseDate = today;
                        trb.CloseContent = "Stop using die";
                        DieStatusUpdateRegular newUpdate = new DieStatusUpdateRegular();
                        newUpdate.DieID = trb.DieID;
                        newUpdate.DieNo = trb.DieNo;
                        newUpdate.RecodeDate = today.ToString("yyyy-MM-dd");
                        newUpdate.ActualShort = trb.ActualShort.Value.ToString();
                        newUpdate.DetailUsingStatus = "Stop using follow TPI: " + trb.TroubleNo;
                        newUpdate.DieOperationStatus = trb.DetailPhenomenon;
                        newUpdate.StopDate = today.ToString("yyyy-MM-dd");
                        newUpdate.DieStatus = "Stop";
                        newUpdate.IssueDate = today;
                        newUpdate.IssueBy = name;
                        db.DieStatusUpdateRegulars.Add(newUpdate);
                        db.SaveChanges();
                    }

                    trb.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + "Cancel TPI (Cancel MR)";
                    db.Entry(trb).State = EntityState.Modified;
                    db.SaveChanges();
                    sendEmailJob.sendEmailTrouble(trb, "PUR", "", "TPI was canceled", EmailSender);
                }
            }
        }

        public int addNewPart(string partNo, string partNam, string modelID, string commonWith)
        {

            // Add Part
            Parts1 newPart = new Parts1();
            var existPartNo = db.Parts1.Where(x => x.PartNo.Contains(partNo.Trim())).FirstOrDefault();
            int NEWPARTID = 0;
            if (existPartNo != null)
            {
                newPart = existPartNo;
            }
            newPart.PartNo = partNo;
            newPart.PartName = partNam;
            newPart.Model = !String.IsNullOrWhiteSpace(modelID) ? db.ModelLists.Find(int.Parse(modelID)).ModelName : newPart.Model;
            newPart.Note = "Common/Family with" + commonWith;
            newPart.Active = true;


            if (existPartNo != null)
            {
                db.Entry(newPart).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                db.Parts1.Add(newPart);
                db.SaveChanges();
            }
            NEWPARTID = newPart.PartID;

            return NEWPARTID;
        }
        public bool addCommonDie(int DieID, string DieNo, int PartID, string PartNo)
        {
            bool status = false;

            var isCommonExit = db.CommonDie1.Where(x => x.DieID == DieID && x.PartID == PartID).FirstOrDefault();
            if (isCommonExit != null)
            {
                isCommonExit.Active = true;
                db.Entry(isCommonExit).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                CommonDie1 newCommon = new CommonDie1();
                newCommon.DieNo = DieNo.Trim().ToUpper();
                newCommon.PartNo = PartNo.Trim().ToUpper();
                newCommon.DieID = DieID;
                newCommon.PartID = PartID;
                newCommon.Active = true;
                db.CommonDie1.Add(newCommon);
                db.SaveChanges();
            }
            status = true;

            return status;
        }
        public string[] CreateModelName(string modelNameInput)
        {
            //1. Tách model
            string[] delimiterChars = { "/", "\\", "-", "," };
            string[] listModel = modelNameInput.Split(delimiterChars, StringSplitOptions.None);
            int modelID = 0;
            var modelPhase = "";
            string modelName = string.Join("/", listModel);
            foreach (string item in listModel)
            {
                if (item.Length > 3)
                {
                    var findModelExist = db.ModelLists.Where(x => x.ModelName.Contains(item)).FirstOrDefault();
                    if (findModelExist != null)
                    {
                        modelID = findModelExist.ModelID;
                        modelPhase = findModelExist.Phase;
                        break;
                    }
                }
            }
            if (modelID == 0) // Tao ModelName moi
            {
                ModelList newModel = new ModelList();
                newModel.ModelName = modelName;
                newModel.Phase = "MT";
                newModel.IssueDate = DateTime.Now;
                db.ModelLists.Add(newModel);
                db.SaveChanges();
                modelID = newModel.ModelID;
                modelPhase = newModel.Phase;
                var msg = "Model Name " + modelName + " vừa được tạo tự đông, " +
                             "Admin hãy vào cập nhât các thông tin về model này để Issue MR chính xác." +
                             " Bắt buộc có thông tin Color/Mono, Nếu không cập nhật quá trình issue MR có thể bị lỗi ";
                sendEmailJob.sendEmailToAdmin(msg);
            }
            string[] result = new string[] { modelID.ToString(), modelName, modelPhase };

            return result;
        }


        //public void updateDieLaunchingControl(MR mR, string userName)
        //{
        //    try
        //    {
        //        var partNo = mR.PartNo.ToUpper().Trim();
        //        var dieNo = mR.Clasification.ToUpper().Trim();
        //        var result = isRenewOrAddOrMT(dieNo);
        //        if ((result.Contains("Renew") || result.Contains("Add")) && mR.Belong.Contains("LBP"))
        //        {
        //            var existDieLaunching = db.Die_Launch_Management.Where(x => x.Part_No == partNo && x.Die_No == dieNo && x.isActive != false).FirstOrDefault();
        //            if (existDieLaunching == null) // Null thi add new
        //            {
        //                Die_Launch_Management newRecord = new Die_Launch_Management();
        //                newRecord.Part_No = partNo.Trim().ToUpper();
        //                newRecord.Part_Name = mR.PartName.Trim();
        //                newRecord.Die_No = dieNo.Trim().ToUpper();
        //                newRecord.Die_ID = partNo + "-" + dieNo + "-001";
        //                newRecord.Process_Code = db.ProcessCodeCalogories.Find(mR.ProcessCodeID).Type;
        //                newRecord.ModelID = mR.ModelID;
        //                newRecord.SupplierID = mR.SupplierID;
        //                newRecord.MC_Size = mR.MCSize;
        //                newRecord.Cav_Qty = mR.CavQty;
        //                newRecord.MR_Request_Date = mR.RequestDate;
        //                newRecord.MR_App_Date = mR.PAEAppDate;
        //                newRecord.Target_OK_Date = null;
        //                newRecord.Decision_Date = DateTime.Now.ToString("yyyy-MM-dd");
        //                if (result == "MT")
        //                {
        //                    newRecord.Select_Supplier_Date = "-";
        //                }
        //                newRecord.isActive = true;
        //                newRecord.isClosed = false;
        //                newRecord.CreateDate = DateTime.Now;
        //                newRecord.His_Update = userName + ": Issue Item";
        //                newRecord.Latest_Update_By = userName;
        //                newRecord.Latest_Update_Date = DateTime.Now;
        //                db.Die_Launch_Management.Add(newRecord);
        //                db.SaveChanges();
        //            }
        //            else // Ko null thi Edit
        //            {
        //                existDieLaunching.Part_Name = mR.PartName.Trim();
        //                existDieLaunching.Die_No = dieNo.Trim().ToUpper();
        //                existDieLaunching.Process_Code = db.ProcessCodeCalogories.Find(mR.ProcessCodeID).Type;
        //                existDieLaunching.ModelID = mR.ModelID;
        //                existDieLaunching.SupplierID = mR.SupplierID;
        //                existDieLaunching.MC_Size = mR.MCSize;
        //                existDieLaunching.Cav_Qty = mR.CavQty;
        //                existDieLaunching.MR_Request_Date = existDieLaunching.MR_Request_Date == null ? mR.RequestDate : existDieLaunching.MR_Request_Date;
        //                existDieLaunching.MR_App_Date = existDieLaunching.MR_App_Date == null ? mR.PAEAppDate : existDieLaunching.MR_App_Date;
        //                db.Entry(existDieLaunching).State = EntityState.Modified;
        //                db.SaveChanges();
        //            }
        //        }
        //    }
        //    catch
        //    {

        //    }
        //}

        public void UpdateTPIStatus(MR mR, string progress)
        {
            var today = DateTime.Now;
            try
            {
                string[] troubleID = mR.TroubleID?.Split(',');
                if (troubleID != null)
                {
                    foreach (var id in troubleID)
                    {
                        var trbl = db.Troubles.Find(int.Parse(id));
                        if (progress == "MRIssue")
                        {
                            trbl.FinalStatusID = 4; // W-MR-issue
                            trbl.MRRquestDate = mR.RequestDate;

                        }
                        if (progress == "MRApp")
                        {
                            trbl.FinalStatusID = 5; // W-PO-Issue
                            trbl.MRRquestDate = mR.RequestDate;
                            trbl.MR_FinAppDate = today;
                        }

                        if (progress == "POIssue")
                        {
                            trbl.FinalStatusID = 6; // W-FA-Result

                            trbl.PODate_PricePUS = mR.PODate;
                        }
                        db.Entry(trbl).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                }
            }
            catch
            {

            }

        }
        public bool EditMrAll(MR mr, string role, string userName, HttpPostedFileBase purAttach)
        {

            bool resultEdit = false;
            //Check MR StatusID = 1 || 2 || 11 (W-Issue || W-PAE-Check || Rejected) or not?
            var mR = db.MRs.Find(mr.MRID);
            if ((mR.StatusID == 1 || mR.StatusID == 2 || mR.StatusID == 11) || role == "Admin")
            {
                // Kiểm tra model nếu là X1-new

                if (mr.TypeID == 1)
                {
                    //Kiem tra model ton tai chua => new chua thi tao new ModelName
                    string[] findModelName = CreateModelName(mR.ModelName);
                    mR.ModelID = Convert.ToInt32(findModelName[0]);
                    mR.ModelName = findModelName[1];
                    mR.Phase = findModelName[2];
                }
                else
                {
                    mR.ModelID = mr.ModelID != null ? mr.ModelID : mR.ModelID;
                }

                mR.Belong = mr.Belong;
                if (role == "Admin")
                {
                    goto next;
                    // Neu Admin chi su thong tin, Ko sua trang thai/ So MR
                }
                else
                {
                    if (mR.StatusID == 11) // Rejected
                    {
                        if (mr.Belong == "LBP")
                        {
                            mR.StatusID = 2; // W-PAE-Check
                        }
                        if (mr.Belong == "CRG")
                        {
                            mR.StatusID = 16; // W-CRG-App

                            var mainMRNO = mR.MRNo.Remove(mR.MRNo.Length - 2, 2);
                            var getVerMR = mR.MRNo.Substring(mR.MRNo.Length - 2, 2);
                            int verNumConvert = Convert.ToInt16(getVerMR);
                            int newVerMR = verNumConvert + 1;
                            var newVerMRString = Convert.ToString(newVerMR);
                            if (newVerMRString.Length == 1)
                            {
                                newVerMRString = "0" + newVerMRString;
                            }
                            mR.MRNo = mainMRNO + newVerMRString;
                        }
                        if (mr.Belong == "PACKING")
                        {
                            mR.StatusID = 15; //W-PAE-App
                        }

                        if (mR.MRNo == null)
                        {
                            var totalMRinThisYear = db.MRs.Where(x => x.MRNo != null && x.IssueDate.Value.Year == DateTime.Now.Year).Count() + 1;
                            mR.MRNo = "MR" + DateTime.Now.ToString("yyMMdd") + "-" + totalMRinThisYear + "-00";
                        }
                        else
                        {
                            var mainMRNO = mR.MRNo.Remove(mR.MRNo.Length - 2, 2);
                            var getVerMR = mR.MRNo.Substring(mR.MRNo.Length - 2, 2);
                            int verNumConvert = Convert.ToInt16(getVerMR);
                            int newVerMR = verNumConvert + 1;
                            var newVerMRString = Convert.ToString(newVerMR);
                            if (newVerMRString.Length == 1)
                            {
                                newVerMRString = "0" + newVerMRString;
                            }
                            mR.MRNo = mainMRNO + newVerMRString;
                        }

                    }
                }

            next:

                // Edit lại MR nếu có thông tin nào được nhập vào, Ko nhập thì giữ nguyên 
                mR.TypeID = mr.TypeID;
                mR.ModelName = mr.ModelName != null ? mr.ModelName : mR.ModelName;
                mR.DieNo = mr.DieNo != null ? mr.DieNo : mR.DieNo;
                mR.RenewForDie = mr.RenewForDie != null ? mr.RenewForDie : mR.RenewForDie;
                mR.PartNo = mr.PartNo != null ? mr.PartNo.ToUpper().Trim() : mR.PartNo.ToUpper().Trim();
                mR.PartName = mr.PartName != null ? mr.PartName : mR.PartName;
                mR.ProcessCodeID = mr.ProcessCodeID != null ? mr.ProcessCodeID : mR.ProcessCodeID;
                mR.Clasification = mr.Clasification != null ? mr.Clasification : mR.Clasification;
                mR.Reason = mr.Reason != null ? mr.Reason : mR.Reason;
                mR.SupplierID = mr.SupplierID != null ? mr.SupplierID : mR.SupplierID;
                mR.SupplierName = db.Suppliers.Find(mR.SupplierID).SupplierName;
                mR.OrderTo = mr.OrderTo != null ? mr.OrderTo : mR.OrderTo;
                mR.DrawHis = mr.DrawHis != null ? mr.DrawHis : mR.DrawHis;
                mR.ECNNo = mr.ECNNo != null ? mr.ECNNo : mR.ECNNo;
                mR.CavQty = mr.CavQty != null ? mr.CavQty : mR.CavQty;
                mR.MCSize = mr.MCSize != null ? mr.MCSize : mR.MCSize;
                mR.BudgetCode = mr.BudgetCode != null ? mr.BudgetCode : mR.BudgetCode;
                mR.IssueNo = mr.IssueNo != null ? mr.IssueNo : mR.IssueNo;
                mR.EstimateCost = mr.EstimateCost != null ? mr.EstimateCost : mR.EstimateCost;
                mR.AppCost = mr.AppCost != null ? mr.AppCost : mR.AppCost;
                mR.ToolingNo = mr.ToolingNo != null ? mr.ToolingNo : mR.ToolingNo;
                mR.PODate = mr.PODate != null ? mr.PODate : mR.PODate;
                mR.PDD = mr.PDD != null ? mr.PDD : mR.PDD;
                //mR.GLAccount = mr.GLAccount != null ? mr.GLAccount : mR.GLAccount;
                //mR.AssetNumber = mr.AssetNumber != null ? mr.AssetNumber : mR.AssetNumber;
                //mR.Location = mr.Location != null ? mr.Location : mR.Location;
                mR.CommonPart = mr.CommonPart != null ? mr.CommonPart : mR.CommonPart;
                mR.Attachment = mr.Attachment != null ? mr.Attachment : mR.Attachment;
                mR.RenewForDie = mr.RenewForDie != null ? mr.RenewForDie : mR.RenewForDie;
                mR.Note = mr.Note != null ? mr.Note : mR.Note;
                mR.DieSpecial = mr.DieSpecial != null ? mr.DieSpecial : mR.DieSpecial;
                mR.PAEComment = mr.PAEComment != null ? mr.PAEComment : mR.PAEComment;
                mR.PLANComment = mr.PLANComment != null ? mr.PLANComment : mR.PLANComment;
                mR.PURComment = mr.PURComment != null ? mr.PURComment : mR.PURComment;
                mR.Unit = mr.Unit != null ? mr.Unit : mR.Unit;
                mR.DisposeDieID = mr.DisposeDieID;
                mR.SucessDieID = mr.SucessDieID;
                mR.SucessPartNo = mr.SucessPartNo;
                mR.SucessPartName = mr.SucessPartName;
                if (!String.IsNullOrEmpty(mr.Note))
                {
                    mR.Note = DateTime.Now.ToString("yyyy/MM/dd ") + userName + " :" + mr.Note;
                }
                if (purAttach != null)
                {
                    string fileName = purAttach.FileName;
                    fileName = "PURAttach_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + fileName;
                    purAttach.SaveAs(Server.MapPath("~/File/MR/" + fileName));

                    mR.PURAttach = fileName;
                }
                // Kiểm tra Unit to exchange tiền tệ
                if (mr.Unit == "USD")
                {
                    mR.ExchangeRate = null;
                    mR.EstimateCostExchangeUSD = mR.EstimateCost;
                }
                else
                {
                    if (mr.Unit == "VND")
                    {
                        mR.ExchangeRate = db.ExchangeRates.ToList().LastOrDefault().RateVNDtoUSD;
                        mR.EstimateCostExchangeUSD = System.Math.Round(Convert.ToDouble(mR.EstimateCost / mR.ExchangeRate), 2);
                    }
                    else
                    if (mr.Unit == "JPY")
                    {
                        mR.ExchangeRate = db.ExchangeRates.ToList().LastOrDefault().RateJPYtoUSD;
                        mR.EstimateCostExchangeUSD = System.Math.Round(Convert.ToDouble(mR.EstimateCost / mR.ExchangeRate), 2);
                    }
                }
                if (mR.Belong != "CRG")
                {
                    var modelColorOrMono = db.ModelLists.Find(mR.ModelID).ModelType;
                    var supplierCode = db.Suppliers.Find(mR.SupplierID).SupplierCode;
                    var first2LeterPartNo = mR.PartNo.Remove(2, mR.PartNo.Length - 2);
                    string[] accInfor = AutoAccFill(mR.TypeID.Value, modelColorOrMono, supplierCode, mR.EstimateCost.Value, mR.Unit, first2LeterPartNo);
                    mR.GLAccount = accInfor[0];
                    mR.Location = accInfor[1];
                    mR.AssetNumber = accInfor[2];

                    // Xu li reason cho truong hop X7 repair và X6 modify và OH
                    if (mR.TypeID == 6 || mR.TypeID == 5 || mR.TypeID == 7)
                    {
                        try
                        {
                            string[] arrReasonList = mR.Reason.Split(',');
                            var reason = "";
                            var troubleID = "";
                            foreach (var item in arrReasonList)
                            {
                                var itemCut = item.Remove(item.Length - 3, 3);
                                if (!String.IsNullOrWhiteSpace(itemCut))
                                {
                                    var trbl = db.Troubles.Where(x => x.TroubleNo.Contains(itemCut) && x.Active != false).FirstOrDefault();
                                    reason = trbl.TroubleName + "," + reason;
                                    troubleID = trbl.TroubleID + "," + troubleID;
                                }

                            }
                            mR.Reason = reason.Remove(reason.Length - 1);
                            mR.TroubleID = troubleID.Remove(troubleID.Length - 1);
                            if (String.IsNullOrWhiteSpace(mR.TroubleID))
                            {
                                ViewBag.err = "You must input TPI No for case X5 || X6 || X7";

                            }
                        }
                        catch
                        {
                            ViewBag.err = "You must input TPI No for case X5 || X6 || X7";

                        }
                    }
                }

                try
                {
                    db.Entry(mR).State = EntityState.Modified;
                    db.SaveChanges();

                    genarateNewDie("", "", "", "", "", "", "", "", "", mR, null);

                    resultEdit = true;
                }
                catch
                {
                    resultEdit = false;
                }

            }

            return resultEdit;
        }

        public int countHoliday(DateTime startDate, DateTime endDate)
        {
            int output = 0;
            int count = dbHolidays.Calendars.Where(x => x.CalendarDate >= startDate && x.CalendarDate <= endDate && x.Holiday == true).Count();
            output = count;
            return output;
        }

        public string getTroubleType(Trouble trbl)
        {
            var type = "";
            if (!String.IsNullOrWhiteSpace(trbl.TroubleTypeID))
            {
                string[] troubeTypeID = trbl.TroubleTypeID?.Split(',');
                foreach (string id in troubeTypeID)
                {
                    type = db.TroubleTypeCalogories.Find(int.Parse(id)).TroubleType + ", " + type;
                }
            }

            return type;
        }
        public string getTroubleRootCause(Trouble trbl)
        {
            var cause = "";
            if (!String.IsNullOrWhiteSpace(trbl.RootCauseID))
            {
                string[] RootCauseID = trbl.RootCauseID?.Split(',');
                foreach (string id in RootCauseID)
                {
                    cause = db.TroubleRootCauseCalogories.Find(int.Parse(id)).RootCause + ", " + cause;
                }
            }
            return cause;
        }
        public string getTempoCM(Trouble trbl)
        {
            var cm = "";
            if (!String.IsNullOrWhiteSpace(trbl.TempCMID))
            {
                string[] tempID = trbl.TempCMID?.Split(',');
                foreach (string id in tempID)
                {
                    cm = db.TempCMcalolgories.Find(int.Parse(id)).TempCM + ", " + cm;
                }
            }
            return cm;
        }
        public string getPerCM(Trouble trbl)
        {
            var cm = "";
            if (!String.IsNullOrWhiteSpace(trbl.PerCMID))
            {
                string[] perID = trbl.PerCMID?.Split(',');
                foreach (string id in perID)
                {
                    cm = db.PerCMCalogories.Find(int.Parse(id)).PerCM + ", " + cm;
                }
            }
            return cm;
        }
        public string getTPINo(string TroublID)
        {
            string output = "";
            if (!String.IsNullOrEmpty(TroublID))
            {
                string[] listID = TroublID.Split(',');
                foreach (var id in listID)
                {
                    int intID = int.Parse(id);
                    output += db.Troubles.Find(intID).TroubleNo;
                }
            }
            return output;
        }
        public JsonResult getListFAResultCategory()
        {
            var result = faResultCategory();

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getListSupplier()
        {
            db.Configuration.ProxyCreationEnabled = false;
            var result = db.Suppliers.ToList();

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getListModel()
        {
            db.Configuration.ProxyCreationEnabled = false;
            var result = db.ModelLists.ToList();

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public IEnumerable faResultCategory()
        {
            IEnumerable FA_Result_list = new[]
           {
                new { value = "OK", show = "OK" },
                new { value = "RS", show = "RS" },
                new { value = "NG", show = "NG" },
                new { value = "NGA", show = "NGA" },
                new { value = "NGB", show = "NGB" },
                new { value = "WMT", show = "WMT" },
                new { value = "MQA", show = "MQA" },
                new { value = "PQA", show = "PQA" },
                new { value = "PE1", show = "PE1" },
                new { value = "PAE", show = "PAE" },
                new { value = "Wait", show = "Wait" },
                new { value = "OQA", show = "OQA" },
            };
            return FA_Result_list;
        }

        public string getStatus(Die1 x)
        {
            // Late : FA OK/RS and Date Result > Target OK
            // Earlier : FA OK/RS and Date Result < Target OK 
            // OnTime : FA OK/RS and Date Result  = target OK
            // OnProgres : FA NG || Null : Today + 20 < target OK
            // Pending(x days) : FA NG || Null : Today > Target OK => x = today -target
            // Warning(x days left) : FA NG|| Null : Today >= target -20 && today < target  => x = target - today
            // Plz Input Target OK: Not input target OK.
            // Plz Input FA Result Date
            var today = DateTime.Now;
            string outPut = "";

            if (x.isCancel == true || x.DieClassify == "MP")
            {
                outPut = "-";
                return outPut;
            }

            if (x.Target_OK_Date == null)
            {
                outPut = "Plz input Target OK";
            }
            else
            {
                if (x.FA_Result == "OK" || x.FA_Result == "RS")
                {
                    if (x.FA_Result_Date != null)
                    {
                        if (x.FA_Result_Date > x.Target_OK_Date)
                        {
                            outPut = "Late";
                        }
                        else
                        {
                            if (x.FA_Result_Date < x.Target_OK_Date)
                            {
                                outPut = "Earlier";
                            }
                            else
                            {
                                outPut = "OnTime";
                            }
                        }
                    }
                    else
                    {
                        outPut = "Plz Input FA OK Date";
                    }
                }
                else // FA NG
                {
                    if (today > x.Target_OK_Date)
                    {
                        outPut = "Pending( " + (today - x.Target_OK_Date.Value).Days + " days)";
                    }
                    else
                    {
                        if (today < x.Target_OK_Date && today.AddDays(20) >= x.Target_OK_Date)
                        {
                            outPut = "Warning(" + (x.Target_OK_Date.Value - today).Days + " days left)";
                        }
                        else
                        {
                            outPut = "On Progress";
                        }
                    }
                }
            }

            return outPut;
        }


        public pending getPendingAndDeptResponse(Die1 x)
        {



            pending outPut = new pending();
            var today = DateTime.Now;
            var configPending = db.DieLaunchingWarningConfigs.ToList();
            if (x.DieClassify == "MP")
            {
                //if (x.Disposal == true)
                //{
                //    outPut.Pending_Status = "Disposed";
                //}
                //else
                //{
                //    outPut.Pending_Status = "In MP";
                //}

                outPut.Pending_Status = x.DieStatusCategory.Type;

                outPut.Dept_Respone = "-";
                outPut.Progress = 100;
                return outPut;
            }
            if (x.isCancel == true)
            {
                outPut.Pending_Status = "Canceled";
                outPut.Dept_Respone = "-";
                outPut.Progress = 0;
                return outPut;
            }
            if (x.isClosed == true)
            {
                outPut.Pending_Status = "Closed";
                outPut.Dept_Respone = "-";
                outPut.Progress = 100;
                return outPut;
            }

            //if (String.IsNullOrWhiteSpace(x.Select_Supplier_Date))
            //{
            //    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-SelectSupplier")).FirstOrDefault().Status;
            //    outPut.Dept_Respone = "PUS";
            //    outPut.Progress = 0;
            //    return outPut;
            //}

            //if (String.IsNullOrWhiteSpace(x.QTN_Sub_Date))
            //{
            //    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-QTN_Sub")).FirstOrDefault().Status;
            //    outPut.Dept_Respone = "PUR";
            //    outPut.Progress = 5;
            //    return outPut;
            //}

            //if (String.IsNullOrWhiteSpace(x.QTN_App_Date))
            //{
            //    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-QTN_App")).FirstOrDefault().Status;
            //    outPut.Dept_Respone = "PUS";
            //    outPut.Progress = 10;
            //    return outPut;
            //}

            if (String.IsNullOrWhiteSpace(x.DFM_Sub_Date))
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-DFM_Sub")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PUR";
                outPut.Progress = 15;
                return outPut;
            }
            if (String.IsNullOrWhiteSpace(x.DFM_PAE_Check_Date))
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-DFM_PAE_Check")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PAE";
                outPut.Progress = 15;
                return outPut;
            }
            if (String.IsNullOrWhiteSpace(x.DFM_PE_Check_Date))
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-DFM_PE1_Check")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PE1";
                outPut.Progress = 15;
                return outPut;
            }
            if (String.IsNullOrWhiteSpace(x.DFM_PE_App_Date))
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-DFM_PE1_App")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PE1";
                outPut.Progress = 20;
                return outPut;
            }
            if (String.IsNullOrWhiteSpace(x.DFM_PAE_App_Date))
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-DFM_PAE_App")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PAE";
                outPut.Progress = 25;
                return outPut;
            }
            if (x.MR_Request_Date == null)
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-MR_Issue")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PUR";
                outPut.Progress = 30;
                return outPut;
            }
            if (x.MR_App_Date == null)
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-MR_App")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PAE";
                outPut.Progress = 30;
                return outPut;
            }
            if (x.PO_Issue_Date == null)
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PO_Issue")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PUR";
                outPut.Progress = 35;
                return outPut;
            }
            if (x.PODate == null)
            {
                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PO_App")).FirstOrDefault().Status;
                outPut.Dept_Respone = "PUS";
                outPut.Progress = 35;

                return outPut;
            }


            if (x.T0_Actual == null && String.IsNullOrWhiteSpace(x.T0_Result))// chưa có T0 Actual || result
            {
                if (x.T0_Plan == null) // chưa có T0 plan
                {
                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-T0_Plan")).FirstOrDefault().Status;
                    outPut.Dept_Respone = "PUR";
                    outPut.Progress = 40;

                    return outPut;
                }
                else // Đã có TO plan
                {
                    if (x.T0_Plan > today)
                    {
                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-T0_trial")).FirstOrDefault().Status;
                        outPut.Dept_Respone = "Supplier";
                        outPut.Progress = 40;
                        return outPut;
                    }
                    else // Đã qua T0 Plan
                    {
                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("PlzConfirmT0Result")).FirstOrDefault().Status;
                        outPut.Dept_Respone = "PAE";
                        outPut.Progress = 45;
                        return outPut;
                    }
                }
            }
            else // Đã có Actual || result To
            {
                // Kiểm tra FA PLan
                if (x.FA_Plan == null) // Chưa có FA plan
                {
                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_Plan")).FirstOrDefault().Status;
                    outPut.Dept_Respone = "PUR";
                    outPut.Progress = 45;

                    return outPut;
                }
                else // Đã có FA Plan
                {
                    if (today < x.FA_Plan) // chưa đến Plan
                    {
                        if (String.IsNullOrWhiteSpace(x.FA_Result)) // chưa có Result
                        {
                            outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_Submit")).FirstOrDefault().Status;
                            outPut.Dept_Respone = "Supplier";
                            outPut.Progress = 45;
                            return outPut;
                        }
                        else // Đã có FA result
                        {
                            if (x.FA_Result == "OK" || x.FA_Result == "RS" || x.FA_Result == "WMT") // FA OK
                            {
                                outPut.Pending_Status = "Done";
                                outPut.Dept_Respone = "-";
                                outPut.Progress = 100;
                                return outPut;
                                //if (x.DieClassify == "MT")
                                //{
                                //    outPut.Pending_Status = "Done";
                                //    outPut.Dept_Respone = "-";
                                //    outPut.Progress = 100;
                                //    return outPut;
                                //}
                                //else // RN & AD & OH
                                //{
                                //    if (String.IsNullOrWhiteSpace(x.TVP_No)) // chưa issue TVP
                                //    {
                                //        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Issue")).FirstOrDefault().Status;
                                //        outPut.Dept_Respone = "PE1";
                                //        outPut.Progress = 80;
                                //        return outPut;
                                //    }
                                //    else // đã issue TVP
                                //    {
                                //        if (String.IsNullOrWhiteSpace(x.TVP_Result)) // chưa có TVP result
                                //        {
                                //            outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Result")).FirstOrDefault().Status; ;
                                //            outPut.Dept_Respone = "PDC";
                                //            outPut.Progress = 90;
                                //            return outPut;
                                //        }
                                //        else // đã có TVP result
                                //        {
                                //            if (x.TVP_Result.ToUpper().Contains("OK") || x.TVP_Result.ToUpper() == "-")
                                //            {
                                //                if (String.IsNullOrWhiteSpace(x.PCAR_Result))
                                //                {
                                //                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PCAR_Result")).FirstOrDefault().Status;
                                //                    outPut.Dept_Respone = "MQA";
                                //                    outPut.Progress = 95;
                                //                    return outPut;
                                //                }
                                //                else
                                //                {
                                //                    if (x.PCAR_Result.ToUpper().Contains("OK") || x.PCAR_Result.ToUpper() == "-")
                                //                    {
                                //                        outPut.Pending_Status = "Done";
                                //                        outPut.Dept_Respone = "-";
                                //                        outPut.Progress = 100;
                                //                        return outPut;
                                //                    }
                                //                    else
                                //                    {
                                //                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PCAR_Result")).FirstOrDefault().Status;
                                //                        outPut.Dept_Respone = "MQA";
                                //                        outPut.Progress = 95;
                                //                        return outPut;
                                //                    }
                                //                }
                                //            }
                                //            else // TVP NG
                                //            {
                                //                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Result")).FirstOrDefault().Status;
                                //                outPut.Dept_Respone = "PDC";
                                //                outPut.Progress = 90;
                                //                return outPut;
                                //            }
                                //        }
                                //    }

                                //}
                            }
                            else // FA NG || đang đánh giá 
                            {
                                if (x.FA_Result == "NG" || x.FA_Result == "NGA" || x.FA_Result == "NGB") // FA NG
                                {
                                    if (String.IsNullOrWhiteSpace(x.FA_Action_Improve)) // PAE chưa instruct
                                    {
                                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-RepairMethod")).FirstOrDefault().Status;
                                        outPut.Dept_Respone = "PAE";
                                        outPut.Progress = 60;
                                        return outPut;
                                    }
                                    else // PAE đã instruct
                                    {
                                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_ReSubmit")).FirstOrDefault().Status;
                                        outPut.Dept_Respone = "Supplier";
                                        outPut.Progress = 70;
                                        return outPut;
                                    }
                                }
                                else // Đang Đánh giá FA
                                {
                                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_Result")).FirstOrDefault().Status;
                                    outPut.Dept_Respone = "PE1";
                                    outPut.Progress = 50;
                                    return outPut;
                                }
                            }
                        }

                    }
                    else // đã đến Plan 
                    {
                        if (String.IsNullOrEmpty(x.FA_Result)) // chưa có FA result 
                        {
                            outPut.Pending_Status = configPending.Where(y => y.Status.Contains("PlzConfirmFASubmit?")).FirstOrDefault().Status;
                            outPut.Dept_Respone = "PUR";
                            outPut.Progress = 50;
                            return outPut;
                        }
                        else // ĐÃ có FA Result 
                        {
                            if (x.FA_Result == "OK" || x.FA_Result == "RS" || x.FA_Result == "WMT") // FA OK
                            {
                                outPut.Pending_Status = "Done";
                                outPut.Dept_Respone = "-";
                                outPut.Progress = 100;
                                return outPut;
                                //if (x.DieClassify == "MT")
                                //{
                                //    outPut.Pending_Status = "Done";
                                //    outPut.Dept_Respone = "-";
                                //    outPut.Progress = 100;
                                //    return outPut;
                                //}
                                //else // RN & AD & OH
                                //{
                                //    if (String.IsNullOrWhiteSpace(x.TVP_No)) // chưa issue TVP
                                //    {
                                //        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Result")).FirstOrDefault().Status;
                                //        outPut.Dept_Respone = "PE1";
                                //        outPut.Progress = 80;
                                //        return outPut;
                                //    }
                                //    else // đã issue TVP
                                //    {
                                //        if (String.IsNullOrWhiteSpace(x.TVP_Result)) // chưa có TVP result
                                //        {
                                //            outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Result")).FirstOrDefault().Status;
                                //            outPut.Dept_Respone = "PDC";
                                //            outPut.Progress = 90;

                                //            return outPut;
                                //        }
                                //        else // đã có TVP result
                                //        {
                                //            if (x.TVP_Result.ToUpper().Contains("OK") || x.TVP_Result.ToUpper() == "-")
                                //            {
                                //                if (String.IsNullOrWhiteSpace(x.PCAR_Result))
                                //                {
                                //                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PCAR_Result")).FirstOrDefault().Status;
                                //                    outPut.Dept_Respone = "MQA";
                                //                    outPut.Progress = 95;
                                //                    return outPut;
                                //                }
                                //                else
                                //                {
                                //                    if (x.PCAR_Result.ToUpper().Contains("OK") || x.PCAR_Result.ToUpper().Contains("-"))
                                //                    {
                                //                        outPut.Pending_Status = "Done";
                                //                        outPut.Dept_Respone = "-";
                                //                        outPut.Progress = 100;
                                //                        return outPut;
                                //                    }
                                //                    else
                                //                    {
                                //                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-PCAR_Result")).FirstOrDefault().Status;
                                //                        outPut.Dept_Respone = "MQA";
                                //                        outPut.Progress = 95;
                                //                        return outPut;
                                //                    }
                                //                }
                                //            }
                                //            else // TVP NG
                                //            {
                                //                outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-TVP_Result")).FirstOrDefault().Status;
                                //                outPut.Dept_Respone = "PDC";
                                //                outPut.Progress = 90;
                                //                return outPut;
                                //            }
                                //        }
                                //    }

                                //}
                            }
                            else // FA NG || Đang đánh giá
                            {
                                if (x.FA_Result == "NG" || x.FA_Result == "NGA" || x.FA_Result == "NGB") // FA NG
                                {
                                    if (String.IsNullOrWhiteSpace(x.FA_Action_Improve)) // PAE chưa instruct
                                    {
                                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-RepairMethod")).FirstOrDefault().Status;
                                        outPut.Dept_Respone = "PAE";
                                        outPut.Progress = 60;
                                        return outPut;
                                    }
                                    else // PAE đã instruct
                                    {
                                        outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_Plan")).FirstOrDefault().Status;
                                        outPut.Dept_Respone = "PUR";
                                        outPut.Progress = 45;
                                        return outPut;
                                    }
                                }
                                else // Đang Đánh giá FA
                                {
                                    outPut.Pending_Status = configPending.Where(y => y.Status.Contains("W-FA_Result")).FirstOrDefault().Status;
                                    outPut.Dept_Respone = "PE1";
                                    outPut.Progress = 50;
                                    return outPut;
                                }
                            }
                        }
                    }
                }
            }
        }

        public string getWarning(Die1 x)
        {
            string outPut = "";
            if (x.DieClassify == "MP")
            {
                outPut = "-";
                return outPut;
            }
            var getOverDL = getOverDeadline(x);
            if (getOverDL.isOver == true)
            {
                outPut = "Over Deadline " + getOverDL.dayOver + " day(s)";
            }
            if (getOverDL.isOver == false && !getOverDL.pendingStatus.Contains("Close") && !getOverDL.pendingStatus.Contains("Cancel") && !getOverDL.pendingStatus.Contains("Done"))
            {
                outPut = ("On Deadline");
            }
            else
            {
                outPut = ("-");
            }
            if (x.Texture == null)
            {
                outPut = outPut + System.Environment.NewLine + ("Lets confirm has Texture or not?");
            }
            if (x.JIG_Using == null)
            {
                outPut = outPut + System.Environment.NewLine + ("Lets confirm has JIG or not?");
            }
            if (x.JIG_Using == true)
            {
                if (String.IsNullOrWhiteSpace(x.JIG_ETA_Supplier))
                {
                    outPut = outPut + System.Environment.NewLine + ("Lets confirm JIG ETA supplier");
                }
                else
                {
                    DateTime JIGETA;
                    var result = DateTime.TryParse(x.JIG_ETA_Supplier, out JIGETA);
                    if (result)
                    {
                        if (JIGETA > x.T0_Plan)
                        {
                            outPut = outPut + System.Environment.NewLine + ("JIG ETA later than T0 plan => plz push up JIG");
                        }
                    }
                }
            }


            return outPut;
        }


        public overDeadline getOverDeadline(Die1 x)
        {
            overDeadline outPut = new overDeadline();
            var pendingStatus = getPendingAndDeptResponse(x);
            var today = DateTime.Now;
            outPut.pendingStatus = pendingStatus.Pending_Status;
            outPut.Dept_Respone = pendingStatus.Dept_Respone;
            if (!pendingStatus.Pending_Status.Contains("Done") && !pendingStatus.Pending_Status.Contains("Cancel") && !pendingStatus.Pending_Status.Contains("Close") && !pendingStatus.Pending_Status.Contains("W-T0_trial") && !pendingStatus.Pending_Status.Contains("W-FA_Submit"))
            {
                var config = db.DieLaunchingWarningConfigs.Where(y => y.Status.Contains(pendingStatus.Pending_Status)).FirstOrDefault();
                try
                {
                    PropertyInfo[] targetProps = x.GetType()
                   .GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty);
                    var targetProp = targetProps.FirstOrDefault(y => y.Name == config.CompareWith);
                    if (targetProp != null)
                    {
                        var targetCompare = targetProp.GetValue(x).ToString();
                        DateTime dateCompare;
                        bool isdate = DateTime.TryParse(targetCompare, out dateCompare);
                        if (isdate)
                        {
                            int holiday_Day = countHoliday(dateCompare, today);
                            int Pending_Day = (today - dateCompare).Days - holiday_Day;
                            // Compaire with leadTime
                            if (x.DieClassify == "MT")
                            {
                                if (Pending_Day >= config.MT_Leadtime)
                                {
                                    outPut.isOver = true;
                                    outPut.dayOver = Pending_Day - (int)config.MT_Leadtime;
                                }
                                else
                                {
                                    outPut.isOver = false;
                                    outPut.dayOver = 0;
                                }
                            }
                            else // MP
                            {
                                if (Pending_Day >= config.MP_Leadtime)
                                {
                                    outPut.isOver = true;
                                    outPut.dayOver = Pending_Day - (int)config.MT_Leadtime;
                                }
                                else
                                {
                                    outPut.isOver = false;
                                    outPut.dayOver = 0;
                                }
                            }
                        }

                    }
                }
                catch
                {

                }

            }
            return outPut;
        }
        public class pending
        {
            public string Pending_Status { set; get; }
            public string Dept_Respone { set; get; }
            public int Progress { set; get; }
        }

        public class overDeadline
        {
            public bool isOver { set; get; }
            public int dayOver { set; get; }
            public string pendingStatus { set; get; }
            public string Dept_Respone { set; get; }
        }
        public class listPending
        {
            public pending pending { set; get; }
        }




        public JsonResult convertMTDie()
        {
            var allRecords = db.Die1.Where(x => x.DieClassify == "MT").ToList();

            List<String> result = new List<string>();
            foreach (var record in allRecords)
            {
                record.isClosed = false;
                record.Decision_Date = "-";
                record.Select_Supplier_Date = "-";
                record.QTN_Sub_Date = "-";
                record.QTN_App_Date = "-";
                record.Need_Use_Date = "-";
                record.DFM_Sub_Date = "-";
                record.DFM_PAE_Check_Date = "-";
                record.DFM_PAE_App_Date = "-";
                record.DFM_PE_Check_Date = "-";
                record.DFM_PE_App_Date = "-";
                record.T0_Result = "-";
                record.FA_Problem = "-";
                record.FA_Action_Improve = "-";
                MR existMR = db.MRs.Where(x => x.PartNo == record.PartNoOriginal && x.Clasification == record.Die_Code && x.Active != false && x.StatusID != 11 && x.StatusID != 12).FirstOrDefault();
                if (existMR != null)
                {
                    record.MR_Request_Date = existMR.RequestDate;
                    record.MR_App_Date = existMR.PURAppDate;
                    PO_Dies existPO = db.PO_Dies.Where(x => x.MRID == existMR.MRID && x.POStatusID != 20 && x.Active != false).FirstOrDefault();
                    if (existPO != null)
                    {
                        record.PO_Issue_Date = existPO.IssueDate;
                        record.PODate = existPO.PODate;
                        record.Target_OK_Date = existPO.PODate.HasValue ? Convert.ToDateTime(existPO.PODate).AddDays(120) : record.Target_OK_Date;
                    }

                }


                db.Entry(record).State = EntityState.Modified;
                db.SaveChanges();
                result.Add(record.DieNo);
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult convertMRTPI()
        {
            List<string> output = new List<string>();
            List<MR> mRs = db.MRs.Where(x => x.TroubleNo_OnlyForX7 != null).ToList();
            foreach (MR mr in mRs)
            {
                Trouble findTPI = db.Troubles.Where(x => x.TroubleNo.Contains(mr.TroubleNo_OnlyForX7.Remove(mr.TroubleNo_OnlyForX7.Length - 3, 3))).FirstOrDefault();
                mr.TroubleID = findTPI?.TroubleID.ToString();
                db.Entry(mr).State = EntityState.Modified;
                db.SaveChanges();
                output.Add(mr.MRNo);

            }

            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public class exchangeCurrency
        {
            public double? rate { set; get; }
            public double? price { set; get; }
        }
        public exchangeCurrency exchangeToUSD(double? price, string currency)
        {
            double? outputPrice = 0.00;
            double? rate = 0;
            if (currency == "USD")
            {
                //mR.ExchangeRate = null;
                outputPrice = price;
                rate = 1;
            }
            else
            {
                if (currency == "VND")
                {
                    rate = db.ExchangeRates.ToList().LastOrDefault().RateVNDtoUSD;
                    outputPrice = System.Math.Round(Convert.ToDouble(price / rate), 2);
                }
                else
                if (currency == "JPY")
                {
                    rate = db.ExchangeRates.ToList().LastOrDefault().RateJPYtoUSD;
                    outputPrice = System.Math.Round(Convert.ToDouble(price / rate), 2);
                }
            }
            exchangeCurrency output = new exchangeCurrency()
            {
                rate = rate,
                price = outputPrice
            };
            return output;
        }

        public void readFilesInventResult(string fullPathfile, string mail, string userName, string dept)
        {
            string msg = "";
            int count = 0;
            try
            {
                var today = DateTime.Now;

                string fileName = Path.GetFileName(fullPathfile);

                Dispose_ControlFileUpload newFile = new Dispose_ControlFileUpload()
                {
                    FileName = fileName,
                    Type = "InventoryResult",
                    Dept = dept,
                    UploadBy = userName,
                    UploadDate = today,
                    Active = false
                };
                db.Dispose_ControlFileUpload.Add(newFile);
                db.SaveChanges();
                using (ExcelPackage package = new ExcelPackage(fullPathfile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    for (int row = start.Row + 5; row <= end.Row; row++)
                    { // Row by row...
                        string assetNo = worksheet.Cells[row, 3].Text.ToUpper().Trim();

                        try
                        {

                            string oldAssetNo = worksheet.Cells[row, 4].Text.ToUpper().Trim();
                            string codeCenter = worksheet.Cells[row, 5].Text.ToUpper().Trim();
                            string deptName = worksheet.Cells[row, 6].Text.ToUpper().Trim();
                            string assetName = worksheet.Cells[row, 7].Text.ToUpper().Trim();
                            string note1 = worksheet.Cells[row, 8].Text.ToUpper().Trim();
                            string component = worksheet.Cells[row, 9].Text.ToUpper().Trim();
                            string classCode = worksheet.Cells[row, 10].Text.ToUpper().Trim();
                            string originCost = worksheet.Cells[row, 11].Text.ToUpper().Trim();
                            string remainCost = worksheet.Cells[row, 12].Text.ToUpper().Trim();
                            string startDate = worksheet.Cells[row, 13].Text.ToUpper().Trim();
                            string location = worksheet.Cells[row, 14].Text.ToUpper().Trim();
                            string dieID = worksheet.Cells[row, 15].Text.ToUpper().Trim();
                            string isFixAssetString = worksheet.Cells[row, 16].Text.ToUpper().Trim();
                            string isFAplate = worksheet.Cells[row, 17].Text.ToUpper().Trim();
                            string usingStatus = worksheet.Cells[row, 18].Text.ToUpper().Trim();
                            string stopDate = worksheet.Cells[row, 19].Text.ToUpper().Trim();
                            string actionPlan = worksheet.Cells[row, 20].Text.ToUpper().Trim();
                            string reasonforDispose = worksheet.Cells[row, 21].Text.ToUpper().Trim();
                            string shot = worksheet.Cells[row, 22].Text.ToUpper().Trim();
                            string RecordshotDate = worksheet.Cells[row, 23].Text.ToUpper().Trim();
                            string isWrongLocation = worksheet.Cells[row, 24].Text.ToUpper().Trim();


                            // verifydata
                            int numComponent = 0;
                            bool isNumberComponent = int.TryParse(component, out numComponent);

                            double dblOriginCost = 0.00;
                            bool isNumberCost = double.TryParse(originCost, out dblOriginCost);

                            double dblRemainCost = 0.00;
                            bool isNumberRemainCost = double.TryParse(remainCost, out dblRemainCost);

                            DateTime dateStartUse;
                            bool isDateStartUse = DateTime.TryParse(startDate, out dateStartUse);

                            bool isFixAsset = isFixAssetString.Contains("YES") || isFixAssetString.Contains("Y") ? true : false;

                            DateTime dateStopUse;
                            bool isDateStopUse = DateTime.TryParse(stopDate, out dateStopUse);

                            double dblshot = 0;
                            bool isNumbershot = double.TryParse(shot, out dblshot);

                            DateTime dateShot;
                            bool isDateShot = DateTime.TryParse(RecordshotDate, out dateShot);

                            // lưu dữ liệu vào database
                            InventoryResult newItemInvent = new InventoryResult()
                            {
                                FileInventID = newFile.FileID,
                                AssetNo = assetNo,
                                OldAssetNo = oldAssetNo,
                                CostCenter = codeCenter,
                                DeptName = deptName,
                                AssetName = assetName,
                                Note1 = note1,
                                ComponentAsset = numComponent,
                                ClassCode = classCode,
                                OriginalCostUSD = dblOriginCost,
                                RemainCostUSD = dblRemainCost,
                                StartUseDate = isDateStartUse == true ? dateStartUse : (DateTime?)null,
                                LocationSupplierCode = location,
                                DieNo = dieID,
                                IsFixAsset = isFixAsset,
                                FAPlate = isFAplate,
                                UsingStatus = usingStatus,
                                StopDate = isDateStopUse == true ? dateStopUse : (DateTime?)null,
                                ActionPlanForUnuse = actionPlan,
                                ReasonforDispose = reasonforDispose,
                                Shot = dblshot,
                                RecordShotDate = isDateShot == true ? dateShot : (DateTime?)null,
                                IsWrongLocation = isWrongLocation.Contains("Y") || isWrongLocation.Contains("YES") ? true : false,
                                Active = true,
                                isReferDie = false
                            };
                            db.InventoryResults.Add(newItemInvent);
                            db.SaveChanges();
                            count++;
                        }
                        catch
                        {
                            msg += "Asset: " + assetNo + "fail to upload" + Environment.NewLine;
                        }

                    }


                }

                newFile.Active = true;
                db.Entry(newFile).State = EntityState.Modified;
                db.SaveChanges();
                // System checking
                // gọi store procedure update thông tin die và update thông tin inventory
                storeProcudure.procudureVerifyInventoryResult(newFile.FileID);
                storeProcudure.procudureUpdateDieDatabaseFollowInventory(newFile.FileID);

                // send email after finished
                sendEmailJob.sendEmailAfterFinishedUploadFileInventoryResult(mail, userName, fileName, newFile.FileID.ToString(), count.ToString());

            }
            catch
            {
                sendEmailJob.sendEmailAfterFailUploadFileInventoryResult(mail, userName, "fail upload file!");
            }

        }

        public InventoryResult autoAddInventoryResult(Die1 die, int fileInventID, string userName)
        {
            InventoryResult newItemInvent = new InventoryResult()
            {
                FileInventID = fileInventID,
                //AssetNo = assetNo,
                //OldAssetNo = oldAssetNo,
                //CostCenter = codeCenter,
                //DeptName = 
                //AssetName = assetName,
                //Note1 = note1,
                //ComponentAsset = numComponent,
                //ClassCode = classCode,
                //OriginalCostUSD = dblOriginCost,
                //RemainCostUSD = dblRemainCost,
                //StartUseDate = isDateStartUse == true ? dateStartUse : (DateTime?)null,
                LocationSupplierCode = db.Suppliers.Find(die.SupplierID).SupplierCode,
                DieNo = die.DieNo,
                //IsFixAsset = isFixAsset,
                //FAPlate = isFAplate,
                //UsingStatus = usingStatus,
                //StopDate = isDateStopUse == true ? dateStopUse : (DateTime?)null,
                //ActionPlanForUnuse = actionPlan,
                //ReasonforDispose = reasonforDispose,
                //Shot = dblshot,
                //RecordShotDate = isDateShot == true ? dateShot : (DateTime?)null,
                //IsWrongLocation = isWrongLocation.Contains("Y") || isWrongLocation.Contains("YES") ? true : false,
                //ControlNumber = controlNumbers,
                //ControlDept = controlDept,
                isAutoAddReferDie = true,
                AutoAddByWhoAndReason = userName + "_" + DateTime.Now.ToString("yyyy/MM/dd: ") + "Added this die to verify dispose",
                Active = true,
                IsSeletedForVerify = false,
                isReferDie = true,

            };
            db.InventoryResults.Add(newItemInvent);
            db.SaveChanges();
            return newItemInvent;
        }
        public class dieCode
        {
            public string originalCode { set; get; }
            public string renewalCode { set; get; }
        }

        public Die1 getOriginalDie(Die1 renewDie)
        {
            List<dieCode> config = new List<dieCode>()
           {
                 new dieCode { originalCode = "11A", renewalCode = "14A" },
                 new dieCode { originalCode = "21A", renewalCode = "24A" },
                 new dieCode { originalCode = "31A", renewalCode = "34A" },
                 new dieCode { originalCode = "41A", renewalCode = "44A" },
                 new dieCode { originalCode = "51A", renewalCode = "54A" },
                 new dieCode { originalCode = "61A", renewalCode = "64A" },
                 new dieCode { originalCode = "71A", renewalCode = "74A" },
                 new dieCode { originalCode = "81A", renewalCode = "84A" },
                 new dieCode { originalCode = "91A", renewalCode = "94A" },
                 new dieCode { originalCode = "A1A", renewalCode = "A4A" },
                 new dieCode { originalCode = "B1A", renewalCode = "B4A" },

                 new dieCode { originalCode = "14A", renewalCode = "14B" },
                 new dieCode { originalCode = "24A", renewalCode = "24B" },
                 new dieCode { originalCode = "34A", renewalCode = "34B" },
                 new dieCode { originalCode = "44A", renewalCode = "44B" },
                 new dieCode { originalCode = "54A", renewalCode = "54B" },
                 new dieCode { originalCode = "64A", renewalCode = "64B" },
                 new dieCode { originalCode = "74A", renewalCode = "74B" },
                 new dieCode { originalCode = "84A", renewalCode = "84B" },
                 new dieCode { originalCode = "94A", renewalCode = "94B" },
                 new dieCode { originalCode = "A4A", renewalCode = "A4B" },
                 new dieCode { originalCode = "B4A", renewalCode = "B4B" },

                  new dieCode { originalCode = "14B", renewalCode = "14C" },
                 new dieCode { originalCode = "24B", renewalCode = "24C" },
                 new dieCode { originalCode = "34B", renewalCode = "34C" },
                 new dieCode { originalCode = "44B", renewalCode = "44C" },
                 new dieCode { originalCode = "54B", renewalCode = "54C" },
                 new dieCode { originalCode = "64B", renewalCode = "64C" },
                 new dieCode { originalCode = "74B", renewalCode = "74C" },
                 new dieCode { originalCode = "84B", renewalCode = "84C" },
                 new dieCode { originalCode = "94B", renewalCode = "94C" },
                 new dieCode { originalCode = "A4B", renewalCode = "A4C" },
                 new dieCode { originalCode = "B4B", renewalCode = "B4C" },

                   new dieCode { originalCode = "14C", renewalCode = "14D" },
                 new dieCode { originalCode = "24C", renewalCode = "24D" },
                 new dieCode { originalCode = "34C", renewalCode = "34D" },
                 new dieCode { originalCode = "44C", renewalCode = "44D" },
                 new dieCode { originalCode = "54C", renewalCode = "54D" },
                 new dieCode { originalCode = "64C", renewalCode = "64D" },
                 new dieCode { originalCode = "74C", renewalCode = "74D" },
                 new dieCode { originalCode = "84C", renewalCode = "84D" },
                 new dieCode { originalCode = "94C", renewalCode = "94D" },
                 new dieCode { originalCode = "A4C", renewalCode = "A4D" },
                 new dieCode { originalCode = "B4C", renewalCode = "B4D" },

                  new dieCode { originalCode = "14D", renewalCode = "14E" },
                 new dieCode { originalCode = "24D", renewalCode = "24E" },
                 new dieCode { originalCode = "34D", renewalCode = "34E" },
                 new dieCode { originalCode = "44D", renewalCode = "44E" },
                 new dieCode { originalCode = "54D", renewalCode = "54E" },
                 new dieCode { originalCode = "64D", renewalCode = "64E" },
                 new dieCode { originalCode = "74D", renewalCode = "74E" },
                 new dieCode { originalCode = "84D", renewalCode = "84E" },
                 new dieCode { originalCode = "94D", renewalCode = "94E" },
                 new dieCode { originalCode = "A4D", renewalCode = "A4E" },
                 new dieCode { originalCode = "B4D", renewalCode = "B4E" },

            };

            string findOriginalCode = "";
            foreach (var item in config)
            {
                if (item.renewalCode == renewDie.Die_Code)
                {
                    findOriginalCode = item.originalCode;
                    break;
                }
            }

            string orignalDieNo = renewDie.DieNo.Substring(0, 12) + findOriginalCode + "-001";
            var die = db.Die1.Where(x => x.DieNo == orignalDieNo && x.Active != false && x.isOfficial != false).FirstOrDefault();
            return die;


        }

        public JsonResult getAuthor()
        {
            var output = new
            {
                Admin = Session["Admin"]?.ToString(),
                AdminDispose = Session["Code"].ToString() == "DISMEET" ? "Admin" : "",
                UserID = Session["UserID"]?.ToString(),
                Dept = Session["Dept"]?.ToString(),
                Grade = Session["Grade"]?.ToString(),
                Trouble_Role = Session["Trouble_Role"]?.ToString(),
                MR_Role = Session["MR_Role"]?.ToString(),
                PO_Role = Session["PO_Role"]?.ToString(),
                Die_Lauch_Role = Session["Die_Lauch_Role"]?.ToString(),
                Lending_Role = Session["Lending_Role"]?.ToString(),
                DSUM_Role = Session["DSUM_Role"]?.ToString(),
                Dispose_Role = Session["Dispose_Role"]?.ToString(),
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }
        public class DMSRole
        {
            public string Dept { set; get; }
            public string Grade { set; get; }
            public string TPIRole { set; get; }
            public string MRRole { set; get; }
            public string PORole { set; get; }
            public string DieLaunchRole { set; get; }
            public string TransferRole { set; get; }
            public string DSUMRole { set; get; }
            public string DisposalRole { set; get; }
        }
        public DMSRole autoSuggestDMSRoleForUser(User u)
        {
            string[] gradeLevel = { "G1", "G2", "G3", "G4", "G5", "G6", "M1", "AGM", "GM" };
            DMSRole[] configRole =
            {
                //PAE
                new DMSRole {Dept = "PAE",Grade = "G1",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PAE",Grade = "G2",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PAE",Grade = "G3",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PAE",Grade = "G4",TPIRole = "Approve", MRRole = "Check", PORole="View", DieLaunchRole = "Edit", TransferRole = "Approve", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PAE",Grade = "G5",TPIRole = "Approve", MRRole = "Check", PORole="View", DieLaunchRole = "Edit", TransferRole = "Approve", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PAE",Grade = "G6",TPIRole = "Approve", MRRole = "Check", PORole="View", DieLaunchRole = "Edit", TransferRole = "Approve", DSUMRole = "Approve", DisposalRole = "Check"},
                new DMSRole {Dept = "PAE",Grade = "M1",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "Edit", TransferRole = "Approve", DSUMRole = "Approve", DisposalRole = "Approve"},

                 //CRG
                new DMSRole {Dept = "CRG",Grade = "G1",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "deny", TransferRole = "Check", DSUMRole = "deny", DisposalRole = "View"},
                new DMSRole {Dept = "CRG",Grade = "G2",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "deny", TransferRole = "Check", DSUMRole = "deny", DisposalRole = "View"},
                new DMSRole {Dept = "CRG",Grade = "G3",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "deny", TransferRole = "Check", DSUMRole = "deny", DisposalRole = "Check"},
                new DMSRole {Dept = "CRG",Grade = "G4",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "deny", TransferRole = "Check", DSUMRole = "deny", DisposalRole = "Check"},
                new DMSRole {Dept = "CRG",Grade = "G5",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "deny", TransferRole = "Check", DSUMRole = "deny", DisposalRole = "Check"},
                new DMSRole {Dept = "CRG",Grade = "G6",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "deny", TransferRole = "Approve", DSUMRole = "deny", DisposalRole = "Check"},
                new DMSRole {Dept = "CRG",Grade = "M1",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "deny", TransferRole = "Approve", DSUMRole = "deny", DisposalRole = "Approve"},
                new DMSRole {Dept = "CRG",Grade = "AGM",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "deny", TransferRole = "Approve", DSUMRole = "deny", DisposalRole = "Approve"},
                new DMSRole {Dept = "CRG",Grade = "GM",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "deny", TransferRole = "Approve", DSUMRole = "deny", DisposalRole = "Approve"},

                  //DMT
                new DMSRole {Dept = "DMT",Grade = "G1",TPIRole = "Issue", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Issue", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "DMT",Grade = "G2",TPIRole = "Issue", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Issue", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "DMT",Grade = "G3",TPIRole = "Issue", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Issue", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "DMT",Grade = "G4",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "DMT",Grade = "G5",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "DMT",Grade = "G6",TPIRole = "Approve", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "Approve", DisposalRole = "Check"},
                new DMSRole {Dept = "DMT",Grade = "M1",TPIRole = "Approve", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "Approve", DisposalRole = "Approve"},

                //PE1
                new DMSRole {Dept = "PE1",Grade = "G1",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PE1",Grade = "G2",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PE1",Grade = "G3",TPIRole = "Check", MRRole = "View", PORole="View", DieLaunchRole = "Edit", TransferRole = "View", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PE1",Grade = "G4",TPIRole = "Approve", MRRole = "View", PORole="View", DieLaunchRole = "Edit", TransferRole = "View", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PE1",Grade = "G5",TPIRole = "Approve", MRRole = "View", PORole="View", DieLaunchRole = "Edit", TransferRole = "View", DSUMRole = "Check", DisposalRole = "Check"},
                new DMSRole {Dept = "PE1",Grade = "G6",TPIRole = "Approve", MRRole = "View", PORole="View", DieLaunchRole = "Edit", TransferRole = "View", DSUMRole = "Approve", DisposalRole = "Check"},
                new DMSRole {Dept = "PE1",Grade = "M1",TPIRole = "Approve", MRRole = "Approve", PORole="View", DieLaunchRole = "Edit", TransferRole = "View", DSUMRole = "Approve", DisposalRole = "Approve"},

                 //PUR
                new DMSRole {Dept = "PUR",Grade = "G1",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "G2",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "G3",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "G4",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "G5",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "G6",TPIRole = "View", MRRole = "Check", PORole="Check", DieLaunchRole = "Edit", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUR",Grade = "M1",TPIRole = "View", MRRole = "Approve", PORole="Approve", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "View", DisposalRole = "Approve"},
                new DMSRole {Dept = "PUR",Grade = "AGM",TPIRole = "View", MRRole = "Approve", PORole="Approve", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "View", DisposalRole = "Approve"},
                new DMSRole {Dept = "PUR",Grade = "GM",TPIRole = "View", MRRole = "Approve", PORole="Approve", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "View", DisposalRole = "Approve"},

                  //PUS
                new DMSRole {Dept = "PUS",Grade = "G1",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PUS",Grade = "G2",TPIRole = "View", MRRole = "View", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PUS",Grade = "G3",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G4",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G5",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G6",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "M1",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "Approve"},

                  //PUC
                new DMSRole {Dept = "PUS",Grade = "G1",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PUS",Grade = "G2",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PUS",Grade = "G3",TPIRole = "View", MRRole = "View", PORole="Check", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G4",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G5",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "G6",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "Check", DSUMRole = "View", DisposalRole = "Check"},
                new DMSRole {Dept = "PUS",Grade = "M1",TPIRole = "View", MRRole = "View", PORole="Approve", DieLaunchRole = "View", TransferRole = "Approve", DSUMRole = "View", DisposalRole = "Approve"},

                  //Plan
                new DMSRole {Dept = "PLAN",Grade = "G1",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "G2",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "G3",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "G4",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "G5",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "G6",TPIRole = "View", MRRole = "Check", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "M1",TPIRole = "View", MRRole = "Approve", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "AGM",TPIRole = "View", MRRole = "Approve", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
                new DMSRole {Dept = "PLAN",Grade = "GM",TPIRole = "View", MRRole = "Approve", PORole="View", DieLaunchRole = "View", TransferRole = "View", DSUMRole = "View", DisposalRole = "View"},
            };


            DMSRole output = new DMSRole()
            {
                Dept = db.Departments.Find(u.DeptID).DeptName,
                Grade = u.Grade,
                TPIRole = "View",
                MRRole = "View",
                PORole = "View",
                DieLaunchRole = "View",
                TransferRole = "View",
                DSUMRole = "View",
                DisposalRole = "View",
            };
            if (u.DeptID == null || String.IsNullOrWhiteSpace(u.Grade))
            {
                return output;
            }
            foreach (var item in configRole)
            {
                if (db.Departments.Find(u.DeptID).DeptName.Contains(item.Dept) && u.Grade.Contains(item.Grade))
                {
                    output.Dept = item.Dept;
                    output.Grade = item.Grade;
                    output.TPIRole = item.TPIRole;
                    output.MRRole = item.MRRole;
                    output.PORole = item.PORole;
                    output.DieLaunchRole = item.DieLaunchRole;
                    output.TransferRole = item.TransferRole;
                    output.DSUMRole = item.DSUMRole;
                    output.DisposalRole = item.DisposalRole;
                }


            }
            return output;
        }


        //public object readPSFile(string pathFile, string partNo)
        //{
        //    pathFile = "D:\\WebDMS\\CurrentDMS\\DMS03\\DMS03\\File\\PS\\PS_Mid_file_Update_20240621_0847.xlsx";
        //    using (ExcelPackage package = new ExcelPackage(pathFile))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
        //        var start = worksheet.Dimension.Start;
        //        var end = worksheet.Dimension.End;
        //        // Lap dong dau tien de lay tieu de
        //        List<string> ListTitle = new List<string>();


        //        // Lap qua tung dong excel

        //    }
        //}



        //public async Task<JsonResult> Sample1GetJsonData()
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        var sheet = package.Workbook.Worksheets.Add("Currencies");
        //        var csvFileInfo = new FileInfo("D:\\WebDMS\\CurrentDMS\\DMS03\\DMS03\\File\\PS\\PS_Mid_file_Update_20240621_0847.xlsx");
        //        var format = new ExcelTextFormat
        //        {
        //            Delimiter = ';',
        //            Culture = CultureInfo.InvariantCulture,
        //            DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
        //        };
        //        var range = await sheet.Cells["A1"].LoadFromTextAsync(csvFileInfo, format);
        //        sheet.Cells[range.Start.Row, 1, range.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
        //        sheet.Cells[range.Start.Row, 2, range.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";
        //        var jsonData = range.ToJson(x => x.AddDataTypesOn = eDataTypeOn.OnColumn);
        //        return Json(jsonData);
        //    }
        //}



    }
}
