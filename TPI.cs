using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using DMS03.Models;
using System.IO;
using PagedList;
using OfficeOpenXml;
using System.Net.Mail;
using System.Threading;
using Spire.Xls;
using Aspose.Cells;
using Aspose.Slides;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
//using static DMS03.Controllers.TroublesController;
using System.Data.Entity.Migrations.Sql;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System.Web.UI;
using Org.BouncyCastle.Asn1.X509;
using System.Windows.Input;
using System.Windows;
using static iTextSharp.awt.geom.Point2D;

namespace DMS03.Controllers
{
    public class TPIController : Controller
    {
        private DMSEntities db = new DMSEntities();
        SendEmailController sendEmailJob = new SendEmailController();
        CommonFunctionController commonFunction = new CommonFunctionController();

        // GET: TPI
        public ActionResult Index(string Search, string waitFor, string fromDate, string toDate, int? page, string msg, string export, string w_pae, string RPA)
        {
            if (Session["UserID"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            if (page == null) page = 1;
            int pageSize = 50;
            int pageNumber = (page ?? 1);
            var dept = Session["Dept"].ToString();
            var listTrouble = db.Troubles.Where(x => x.Active != false).ToList();

            //For Search
            if (!String.IsNullOrEmpty(Search) || !String.IsNullOrEmpty(fromDate) || !String.IsNullOrEmpty(toDate))
            {
                var searchResult = listTrouble;
                // search text
                if (!String.IsNullOrEmpty(Search))
                {
                    Search = Search.Trim();
                    searchResult = listTrouble.Where(x => x.DieNo != null ? x.DieNo.Contains(Search) : x.TroubleID == 0).ToList();
                    if (searchResult.Count() == 0)
                    {
                        searchResult = listTrouble.Where(x => x.Die1.SupplierID > 0 ? x.Die1.Supplier.SupplierName.Contains(Search) : x.TroubleID == 0).ToList();
                        if (searchResult.Count() == 0)
                        {
                            searchResult = listTrouble.Where(x => x.TroubleNo != null ? x.TroubleNo.Contains(Search) : x.TroubleID == 0).ToList();
                        }
                        if (searchResult.Count() == 0)
                        {
                            var searchResultPIC = listTrouble.Where(x => x.PIC != null ? x.SubmitBy.Contains(Search) : x.TroubleID == 0).ToList();
                            var searchResultSubmitor = listTrouble.Where(x => x.SubmitBy != null ? x.SubmitBy.Contains(Search) : x.TroubleID == 0).ToList();
                            searchResult.AddRange(searchResultSubmitor);
                            searchResult.AddRange(searchResultPIC);
                        }

                        if (searchResult.Count() == 0)
                        {
                            searchResult = listTrouble.Where(x => x.FinalStatusID != null ? x.FinalStatusCalogory.FinalStatus.Contains(Search) : x.TroubleID == 0).ToList();
                        }
                    }
                }
                // Search From 
                if (!String.IsNullOrEmpty(fromDate))
                {
                    searchResult = searchResult.Where(x => x.SubmitDate >= Convert.ToDateTime(fromDate)).ToList();
                }
                if (!String.IsNullOrEmpty(toDate))
                {
                    searchResult = searchResult.Where(x => x.SubmitDate <= Convert.ToDateTime(toDate)).ToList();
                }

                listTrouble = searchResult;
            }

            if (!String.IsNullOrEmpty(RPA))
            {
                listTrouble = db.Troubles.Where(x => x.isNeedRPA == true && x.Active != false).ToList();
                goto ExitSearchWaitFor;
            }


            if (!String.IsNullOrEmpty(waitFor))
            {

                if (waitFor == "PAE-Check_MT")
                {
                    listTrouble = listTrouble.Where(x => x.FinalStatusCalogory.FinalStatus.ToLower().Contains("pae-check") && x.Phase.Contains("MT")).ToList();
                    goto ExitSearchWaitFor;
                }
                if (waitFor == "PAE-App_MT")
                {
                    listTrouble = listTrouble.Where(x => x.FinalStatusCalogory.FinalStatus.ToLower().Contains("pae-app") && x.Phase.Contains("MT")).ToList();
                    goto ExitSearchWaitFor;
                }
                if (waitFor == "PAE-Check_MP")
                {
                    listTrouble = listTrouble.Where(x => x.FinalStatusCalogory.FinalStatus.ToLower().Contains("pae-check") && x.Phase.Contains("MP")).ToList();
                    goto ExitSearchWaitFor;
                }
                if (waitFor == "PAE-App_MP")
                {
                    listTrouble = listTrouble.Where(x => x.FinalStatusCalogory.FinalStatus.ToLower().Contains("pae-app") && x.Phase.Contains("MP")).ToList();
                    goto ExitSearchWaitFor;
                }
                if (waitFor == "W-PUR-Check")
                {
                    listTrouble = listTrouble.Where(x => x.IsNeedPURConfirm == true && x.FinalStatusID != 15).ToList();
                    goto ExitSearchWaitFor;
                }
                listTrouble = listTrouble.Where(x => x.FinalStatusCalogory.FinalStatus.ToLower().Contains(waitFor.ToLower().Trim())).ToList();

            }
        ExitSearchWaitFor:

            if (Session["Dept"].ToString().Contains("CRG"))
            {
                listTrouble = listTrouble.Where(x => x.TroubleFrom.Contains("CRG")).ToList();
            }
            else
            {
                if (Session["Dept"].ToString().Contains("PUR") || Session["Admin"].ToString() == "Admin")
                {
                    listTrouble = listTrouble.ToList();
                }
                else
                {
                    listTrouble = listTrouble.Where(x => !x.TroubleFrom.Contains("CRG")).ToList();
                }
            }


            if (export == "Export")
            {
                ExportExcel(listTrouble);
            }

            ViewBag.msg = msg;
            ViewBag.Search = Search;
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;
            var pendingTrouble = db.Troubles.Where(x => x.FinalStatusID == 1 || x.FinalStatusID == 3 || x.FinalStatusID == 4 || x.FinalStatusID == 5
            || x.FinalStatusID == 6 || x.FinalStatusID == 7 || x.FinalStatusID == 8 || x.FinalStatusID == 9 ||
            x.FinalStatusID == 10 || x.FinalStatusID == 12 || x.FinalStatusID == 13 || x.FinalStatusID == 14).ToList();

            // For CGR
            ViewBag.w_CRGCheck = pendingTrouble.Where(x => x.FinalStatusID == 12 && x.Active == true).Count();
            ViewBag.w_CRGApp = pendingTrouble.Where(x => x.FinalStatusID == 13 && x.Active == true).Count();

            // For LBP
            ViewBag.w_PUR_Check = pendingTrouble.Where(x => x.IsNeedPURConfirm == true && x.Active == true).Count();
            ViewBag.w_MR = pendingTrouble.Where(x => x.FinalStatusID == 4 && x.Active == true).Count();
            ViewBag.w_PO = pendingTrouble.Where(x => x.FinalStatusID == 5 && x.Active == true).Count();
            ViewBag.w_FA = pendingTrouble.Where(x => x.FinalStatusID == 6 && x.Active != false).Count();
            ViewBag.w_DMTCheck = pendingTrouble.Where(x => x.FinalStatusID == 10 && x.Active == true).Count();
            ViewBag.w_DMTApp = pendingTrouble.Where(x => x.FinalStatusID == 7 && x.Active == true).Count();

            ViewBag.w_PE1Check = pendingTrouble.Where(x => x.FinalStatusID == 8 && x.Active == true).Count();
            ViewBag.w_PE1App = pendingTrouble.Where(x => x.FinalStatusID == 14 && x.Active == true).Count();
            ViewBag.w_PAE_C_MT = pendingTrouble.Where(x => x.FinalStatusID == 3 && x.Phase.Contains("MT") && x.Active == true).Count();
            ViewBag.w_PAE_App_MT = pendingTrouble.Where(x => x.FinalStatusID == 9 && x.Phase.Contains("MT") && x.Active == true).Count();
            ViewBag.w_PAE_C_MP = pendingTrouble.Where(x => x.FinalStatusID == 3 && x.Phase.Contains("MP") && x.Active == true).Count();
            ViewBag.w_PAE_App_MP = pendingTrouble.Where(x => x.FinalStatusID == 9 && x.Phase.Contains("MP") && x.Active == true).Count();
            return View(listTrouble.OrderByDescending(x => x.SubmitDate).ToPagedList(pageNumber, pageSize));
        }

        public void sendMailToRPA(int id)
        {
            var trouble = db.Troubles.Find(id);
            sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trouble, null);
        }

        public ActionResult issueTPI()
        {
            if (Session["UserID"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }


            ViewBag.DieID = new SelectList(db.Die1, "DieID", "DieNo");
            ViewBag.DieTroubleRunningStatusID = new SelectList(db.DieTroubleRunningCalogories, "DieTroubleRunningStatusID", "DieTroubelRunningStatus");
            ViewBag.FieldID = new SelectList(db.FieldCalogories, "FieldID", "FieldType");
            ViewBag.NeedDiePO_X6_X7ID = new SelectList(db.NeedDiePOCalogories, "NeedPOID", "PODieCode");
            ViewBag.PerCMID = new SelectList(db.PerCMCalogories, "PerCMID", "PerCM");
            ViewBag.RenewDecisionID = new SelectList(db.RenewDecisionCalogories, "RenewDecisionID", "DecisionStatus");
            ViewBag.ResponsibilityID = new SelectList(db.ResponsibilityCalogories, "ResponsibilityID", "Responsibility");
            ViewBag.TempCMID = new SelectList(db.TempCMcalolgories, "TempCMID", "TempCM");
            ViewBag.RootCauseID = new SelectList(db.TroubleRootCauseCalogories, "RootCauseID", "RootCause");
            ViewBag.TroubleAreaID = new SelectList(db.TroubleAreaCalogories, "TroubleAreaID", "TroubleArea");
            ViewBag.TroubleTypeID = new SelectList(db.TroubleTypeCalogories, "TroubleTypeID", "TroubleType");
            ViewBag.TroubleLevelID = new SelectList(db.TroubleLevelCalogories, "TroubleLevelID", "LevelType");
            return View();

        }





        // Dành cho RPA =>>> Start
        public ActionResult RPAissueTPI_automationFromBOX()
        {

            ViewBag.DieID = new SelectList(db.Die1, "DieID", "DieNo");
            ViewBag.DieTroubleRunningStatusID = new SelectList(db.DieTroubleRunningCalogories, "DieTroubleRunningStatusID", "DieTroubelRunningStatus");
            ViewBag.FieldID = new SelectList(db.FieldCalogories, "FieldID", "FieldType");
            ViewBag.NeedDiePO_X6_X7ID = new SelectList(db.NeedDiePOCalogories, "NeedPOID", "PODieCode");
            ViewBag.PerCMID = new SelectList(db.PerCMCalogories, "PerCMID", "PerCM");
            ViewBag.RenewDecisionID = new SelectList(db.RenewDecisionCalogories, "RenewDecisionID", "DecisionStatus");
            ViewBag.ResponsibilityID = new SelectList(db.ResponsibilityCalogories, "ResponsibilityID", "Responsibility");
            ViewBag.TempCMID = new SelectList(db.TempCMcalolgories, "TempCMID", "TempCM");
            ViewBag.RootCauseID = new SelectList(db.TroubleRootCauseCalogories, "RootCauseID", "RootCause");
            ViewBag.TroubleAreaID = new SelectList(db.TroubleAreaCalogories, "TroubleAreaID", "TroubleArea");
            ViewBag.TroubleTypeID = new SelectList(db.TroubleTypeCalogories, "TroubleTypeID", "TroubleType");
            ViewBag.TroubleLevelID = new SelectList(db.TroubleLevelCalogories, "TroubleLevelID", "LevelType");
            return View();

        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult RPAissueTPI_automationFromBOX(Trouble trouble, HttpPostedFileBase reportFile)
        {

            var err = "";
            if (reportFile == null)
            {
                err = "Please upload file!";
                goto Exit;
            }
            var today = DateTime.Now;
            Session["Dept"] = "PAE";
            Session["Name"] = "RPA";
            Session["Trouble_Role"] = "Issue";
            var dept = Session["Dept"].ToString();

            var sender = "lbp-pae21@canon-vn.com.vn";
            var userName = "RPA";
            bool isReject = false;
            if (ModelState.IsValid)
            {
                // Verify TPI form
                // kết quả isPass = true => xử lí tiếp
                // kết quả isPass = false => Trả về  view (nếu do RPA => reject)
                verify verifyTPIForm = verifyTPIFormat(reportFile, null);
                if (verifyTPIForm.isPass == true)
                {
                    // Kiem tra neu issuer la DMT/MO
                    //if (dept.Contains("DMT") || dept.Contains("MO"))
                    //{

                    //    if (trouble.DetailPhenomenon != null && trouble.DetailAction != null && trouble.TroubleTypeID != null && trouble.TroubleAreaID != null
                    //            && trouble.FeedbackForR_D != null && trouble.FeedbackForDieSpec != null && trouble.TempCMID != null
                    //            && trouble.PerCMID != null && trouble.RootCauseID != null && trouble.NeedDiePO_X6_X7ID != null
                    //            && trouble.RenewDecisionID != null && trouble.NeedPEConfirm != null)
                    //    {
                    //        // Ko có trường nào null
                    //        goto continuse;

                    //    }
                    //    else
                    //    {
                    //        err = "Hãy nhập đủ thông tin ở bảng Sumarize trouble!";
                    //        goto Exit;
                    //    }
                    //}

                }
                else
                {
                    if (userName.Contains("RPA")) // do RPA upload lên
                    {
                        // Reject
                        isReject = true;
                        err = verifyTPIForm.msg;
                        goto continuse;
                    }
                    else
                    {
                        err = verifyTPIForm.msg;
                        goto Exit;
                    }
                }

                err = verifyTPIForm.msg;
            }

        continuse:
            // 1. Doc file va lay thong tin tu TPI
            trouble = readTPI(trouble, reportFile);

            if (isReject == true)
            {
                trouble.FinalStatusID = 11; // Reject
                trouble.TroubleFrom = "OUTSIDE";
                trouble.isNeedRPA = true;

            }
            else
            {
                // 2. Thay doi status
                if (dept == "DMT" || dept == "MO")
                {
                    trouble.TroubleFrom = "INHOUSE";
                    trouble.FinalStatusID = 10; // W-DMT-Check
                }
                else
                {
                    if (db.Die1.Find(trouble.DieID).Belong.Contains("CRG"))
                    {
                        trouble.TroubleFrom = "CRG";
                        trouble.FinalStatusID = 12; // W-CRG-Confirm
                    }
                    else
                    {
                        trouble.TroubleFrom = "OUTSIDE";
                        trouble.FinalStatusID = 3; // W-PAE-Confirm
                    }

                    if (!trouble.TroubleFrom.Contains("IN") && trouble.Phase.Contains("MP") && (trouble.TroubleName.Contains("ECN/ERI") || trouble.TroubleName.Contains("FA improvement") || trouble.TroubleName.Contains("Die transfer") || trouble.TroubleName.Contains("MP trouble")))
                    {

                        if (trouble.PURCommentDate != null)
                        {
                            trouble.IsNeedPURConfirm = false;
                        }
                        else
                        {
                            trouble.IsNeedPURConfirm = true;
                        }
                    }
                    else
                    {
                        trouble.IsNeedPURConfirm = false;
                    }
                }
            }



            trouble.SubmitDate = DateTime.Now;
            trouble.SubmitBy = "RPA";
            trouble.Progress = today.ToString("yyyy-MM-dd") + ": " + "RPA" + " send report.";
            trouble.Active = true;

            if (trouble.SubmitType.Contains("New_Submit") || String.IsNullOrEmpty(trouble.SubmitType) || (trouble.SubmitType == "Revise/Feedback" && String.IsNullOrWhiteSpace(trouble.TroubleNo)))
            {
                trouble.TroubleNo = genarateTPINo(trouble, true);
                trouble = SaveAndUpdateTPIForm(trouble, reportFile, "Submit", err, isReject, false);
                db.Troubles.Add(trouble);
                db.SaveChanges();
            }
            else
            {
                trouble.TroubleNo = genarateTPINo(trouble, false);
                trouble = SaveAndUpdateTPIForm(trouble, reportFile, "Submit", err, isReject, false);
                db.Entry(trouble).State = EntityState.Modified;
                db.SaveChanges();
            }
            convertToPDF(trouble.ReportFromPur, trouble.TroubleID);
            sendEmailJob.anounceNewTPI(trouble);
            if (isReject == true)
            {
                sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trouble, null);
            }
            ViewBag.err = "Success";
            return View();
        //***************************************

        Exit:
            ViewBag.err = err;
            ViewBag.NeedDiePO_X6_X7ID = new SelectList(db.NeedDiePOCalogories, "NeedPOID", "PODieCode");
            ViewBag.PerCMID = new MultiSelectList(db.PerCMCalogories, "PerCMID", "PerCM");
            ViewBag.RenewDecisionID = new SelectList(db.RenewDecisionCalogories, "RenewDecisionID", "DecisionStatus");
            ViewBag.ResponsibilityID = new SelectList(db.ResponsibilityCalogories, "ResponsibilityID", "Responsibility");
            ViewBag.TempCMID = new MultiSelectList(db.TempCMcalolgories, "TempCMID", "TempCM");
            ViewBag.RootCauseID = new MultiSelectList(db.TroubleRootCauseCalogories, "RootCauseID", "RootCause");
            ViewBag.TroubleAreaID = new MultiSelectList(db.TroubleAreaCalogories, "TroubleAreaID", "TroubleArea");
            //ViewBag.PendingProcessID = new SelectList(db.TroublePendingProcessCalogories, "PendingProcessID", "PendingProcess");
            //ViewBag.ProgressID = new SelectList(db.TroubleProgressCalogories, "ProgressID", "Progress");
            ViewBag.TroubleTypeID = new MultiSelectList(db.TroubleTypeCalogories, "TroubleTypeID", "TroubleType");
            return View();
        }
        // Dành cho RPA =>>> End



        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult issueTPI(Trouble trouble, HttpPostedFileBase reportFile)
        {
            if (Session["UserID"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            var err = "";
            if (reportFile == null)
            {
                err = "Please upload file!";
                goto Exit;
            }
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString();

            var sender = Session["Mail"].ToString();
            var userName = Session["Name"].ToString();
            bool isReject = false;
            if (ModelState.IsValid)
            {
                // Verify TPI form
                // kết quả isPass = true => xử lí tiếp
                // kết quả isPass = false => Trả về  view (nếu do RPA => reject)
                verify verifyTPIForm = verifyTPIFormat(reportFile, null);
                if (verifyTPIForm.isPass == true)
                {
                    // Kiem tra neu issuer la DMT/MO
                    //if (dept.Contains("DMT") || dept.Contains("MO"))
                    //{

                    //    if (trouble.DetailPhenomenon != null && trouble.DetailAction != null && trouble.TroubleTypeID != null && trouble.TroubleAreaID != null
                    //            && trouble.FeedbackForR_D != null && trouble.FeedbackForDieSpec != null && trouble.TempCMID != null
                    //            && trouble.PerCMID != null && trouble.RootCauseID != null && trouble.NeedDiePO_X6_X7ID != null
                    //            && trouble.RenewDecisionID != null && trouble.NeedPEConfirm != null)
                    //    {
                    //        // Ko có trường nào null
                    //        goto continuse;

                    //    }
                    //    else
                    //    {
                    //        err = "Hãy nhập đủ thông tin ở bảng Sumarize trouble!";
                    //        goto Exit;
                    //    }
                    //}

                }
                else
                {
                    if (userName.Contains("RPA")) // do RPA upload lên
                    {
                        // Reject
                        isReject = true;
                        err = verifyTPIForm.msg;
                        goto continuse;
                    }
                    else
                    {
                        err = verifyTPIForm.msg;
                        goto Exit;
                    }
                }

                err = verifyTPIForm.msg;
            }

        continuse:
            // 1. Doc file va lay thong tin tu TPI
            trouble = readTPI(trouble, reportFile);

            if (isReject == true)
            {
                trouble.FinalStatusID = 11; // Reject
                trouble.TroubleFrom = "OUTSIDE";
                trouble.isNeedRPA = true;

            }
            else
            {
                // 2. Thay doi status
                if (dept == "DMT" || dept == "MO")
                {
                    trouble.TroubleFrom = "INHOUSE";
                    trouble.FinalStatusID = 10; // W-DMT-Check
                }
                else
                {
                    if (db.Die1.Find(trouble.DieID).Belong.Contains("CRG"))
                    {
                        trouble.TroubleFrom = "CRG";
                        trouble.FinalStatusID = 12; // W-CRG-Confirm
                    }
                    else
                    {
                        trouble.TroubleFrom = "OUTSIDE";
                        trouble.FinalStatusID = 3; // W-PAE-Confirm
                    }

                    if (!trouble.TroubleFrom.Contains("IN") && trouble.Phase.Contains("MP") && (trouble.TroubleName.Contains("ECN/ERI") || trouble.TroubleName.Contains("FA improvement") || trouble.TroubleName.Contains("Die transfer") || trouble.TroubleName.Contains("MP trouble")))
                    {

                        if (trouble.PURCommentDate != null)
                        {
                            trouble.IsNeedPURConfirm = false;
                        }
                        else
                        {
                            trouble.IsNeedPURConfirm = true;
                        }
                    }
                    else
                    {
                        trouble.IsNeedPURConfirm = false;
                    }
                }
            }



            trouble.SubmitDate = DateTime.Now;
            trouble.SubmitBy = Session["Name"].ToString();
            trouble.Progress = today.ToString("yyyy-MM-dd") + ": " + dept + " send report.";
            trouble.Active = true;

            if (trouble.SubmitType.Contains("New_Submit") || (trouble.SubmitType == "Revise/Feedback" && String.IsNullOrWhiteSpace(trouble.TroubleNo)))
            {
                trouble.TroubleNo = genarateTPINo(trouble, true);
                trouble = SaveAndUpdateTPIForm(trouble, reportFile, "Submit", err, isReject, false);
                db.Troubles.Add(trouble);
                db.SaveChanges();
            }
            else
            {
                trouble.TroubleNo = genarateTPINo(trouble, false);
                trouble = SaveAndUpdateTPIForm(trouble, reportFile, "Submit", err, isReject, false);
                db.Entry(trouble).State = EntityState.Modified;
                db.SaveChanges();
            }
            convertToPDF(trouble.ReportFromPur, trouble.TroubleID);
            sendEmailJob.anounceNewTPI(trouble);
            if (isReject == true)
            {
                sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trouble, null);
            }
            return Redirect("/TPI");
        //***************************************

        Exit:
            ViewBag.err = err;
            ViewBag.NeedDiePO_X6_X7ID = new SelectList(db.NeedDiePOCalogories, "NeedPOID", "PODieCode");
            ViewBag.PerCMID = new MultiSelectList(db.PerCMCalogories, "PerCMID", "PerCM");
            ViewBag.RenewDecisionID = new SelectList(db.RenewDecisionCalogories, "RenewDecisionID", "DecisionStatus");
            ViewBag.ResponsibilityID = new SelectList(db.ResponsibilityCalogories, "ResponsibilityID", "Responsibility");
            ViewBag.TempCMID = new MultiSelectList(db.TempCMcalolgories, "TempCMID", "TempCM");
            ViewBag.RootCauseID = new MultiSelectList(db.TroubleRootCauseCalogories, "RootCauseID", "RootCause");
            ViewBag.TroubleAreaID = new MultiSelectList(db.TroubleAreaCalogories, "TroubleAreaID", "TroubleArea");
            //ViewBag.PendingProcessID = new SelectList(db.TroublePendingProcessCalogories, "PendingProcessID", "PendingProcess");
            //ViewBag.ProgressID = new SelectList(db.TroubleProgressCalogories, "ProgressID", "Progress");
            ViewBag.TroubleTypeID = new MultiSelectList(db.TroubleTypeCalogories, "TroubleTypeID", "TroubleType");
            return View();
        }



        public ActionResult detailTPI(int id)
        {
            if (Session["UserID"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }

            var trbl = db.Troubles.Find(id);
            if (trbl.IsHandlePDF != true)
            {
                if (!String.IsNullOrEmpty(trbl.Report))
                {
                    string[] mainName = trbl.Report.Split('.');
                    string newName = mainName[0] + ".pdf";
                    PDFEdit(newName);
                    System.IO.File.Delete(Server.MapPath("~/File/TroubleReport/" + newName));
                }
                if (!String.IsNullOrEmpty(trbl.ReportFromPur))
                {
                    string[] mainName = trbl.ReportFromPur.Split('.');
                    string newName = mainName[0] + ".pdf";
                    PDFEdit(newName);
                    System.IO.File.Delete(Server.MapPath("~/File/TroubleReport/" + newName));
                }
                trbl.IsHandlePDF = true;
                db.Entry(trbl).State = EntityState.Modified;
                db.SaveChanges();
            }
            // Trouble history

            ViewBag.history = db.Troubles.Where(x => x.DieID == trbl.DieID && x.Active != false).OrderByDescending(x => x.SubmitDate).ToList();

            ViewBag.NeedDiePO_X6_X7ID = new SelectList(db.NeedDiePOCalogories, "NeedPOID", "PODieCode", trbl.NeedDiePO_X6_X7ID);
            ViewBag.PerCMID = new MultiSelectList(db.PerCMCalogories, "PerCMID", "PerCM", trbl.PerCMID?.Split(','));
            ViewBag.RenewDecisionID = new SelectList(db.RenewDecisionCalogories, "RenewDecisionID", "DecisionStatus", trbl.RenewDecisionID);
            ViewBag.ResponsibilityID = new SelectList(db.ResponsibilityCalogories, "ResponsibilityID", "Responsibility", trbl.ResponsibilityID);
            ViewBag.TempCMID = new MultiSelectList(db.TempCMcalolgories, "TempCMID", "TempCM", trbl.TempCMID?.Split(','));
            ViewBag.RootCauseID = new MultiSelectList(db.TroubleRootCauseCalogories, "RootCauseID", "RootCause", trbl.RootCauseID?.Split(','));
            ViewBag.TroubleAreaID = new MultiSelectList(db.TroubleAreaCalogories, "TroubleAreaID", "TroubleArea", trbl.TroubleAreaID?.Split(','));
            //ViewBag.PendingProcessID = new SelectList(db.TroublePendingProcessCalogories, "PendingProcessID", "PendingProcess");
            //ViewBag.ProgressID = new SelectList(db.TroubleProgressCalogories, "ProgressID", "Progress");
            ViewBag.TroubleTypeID = new MultiSelectList(db.TroubleTypeCalogories, "TroubleTypeID", "TroubleType", trbl.TroubleTypeID?.Split(','));
            return View(trbl);


        }

        public JsonResult getTotalDie(int troublID)
        {
            var trb = db.Troubles.Find(troublID);
            List<CommonDie1> allDie = db.CommonDie1.Where(x => x.PartNo == trb.PartNo && x.Active != false && x.Die1.isCancel != true && x.Die1.Disposal != true).ToList();
            var dies = allDie.Select(y => new
            {
                DieNo = y.DieNo,
                Status = y.Die1.DieStatusUpdateRegulars.LastOrDefault() != null ? y.Die1.DieStatusUpdateRegulars.LastOrDefault().DieStatus : "-",
                Shot = y.Die1.DieStatusUpdateRegulars.LastOrDefault() != null ? y.Die1.DieStatusUpdateRegulars.LastOrDefault().ActualShort : "-",
                Disposed = y.Die1.Disposal == true ? "Y" : "N",
            });

            return Json(dies, JsonRequestBehavior.AllowGet);
        }


        public JsonResult RejectTPI(int troublID, string rejectReason)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Trouble_Role"].ToString();
            var trbl = db.Troubles.Find(troublID);
            bool status = false;
            string msg = "";
            var sender = Session["Mail"].ToString();
            var today = DateTime.Now;
            var rejectTo = "";
            if (Session["Admin"].ToString() == "Admin")
            {
                // xử lí dữ liệu
                var isTPIIssueMR = db.MRs.Where(x => x.TroubleID != null && x.Active != false && x.StatusID == 1).Where(x => x.TroubleID.Contains(trbl.TroubleID.ToString())).FirstOrDefault();
                if (isTPIIssueMR == null)
                {
                    trbl.FinalStatusID = 11; //Rejected
                    trbl.IsNeedPURConfirm = false;
                    trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                    trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                    rejectTo = "PUR";
                    status = true;
                }
                else
                {
                    status = false;
                    msg = "This TPI already issue MR, if you want to reject, please inform PUR cancel MR first!";
                }

            }
            if (dept.Contains("PAE") && (trbl.FinalStatusID == 3)) // W-PAE-C && App
            {
                // xử lí dữ liệu
                trbl.FinalStatusID = 11; //Rejected
                trbl.IsNeedPURConfirm = false;
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                rejectTo = "PUR";
                status = true;
            }
            if (dept.Contains("CRG") && (trbl.FinalStatusID == 12)) // W-CRG-C && App
            {
                // xử lí dữ liệu
                trbl.FinalStatusID = 11; //Rejected
                trbl.IsNeedPURConfirm = false;
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                rejectTo = "PUR";
                status = true;
            }
            if (dept.Contains("PAE") && role == "Approve" && (trbl.FinalStatusID == 9)) // W-PAE-App
            {
                // xử lí dữ liệu
                trbl.FinalStatusID = 3; //W-PAE-Check
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                rejectTo = "PAE";
                status = true;
            }
            if (dept.Contains("CRG") && role == "Approve" && (trbl.FinalStatusID == 13)) // W-CRG-App
            {
                // xử lí dữ liệu
                trbl.FinalStatusID = 12; //W-CRG-Check
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                rejectTo = "CRG";
                status = true;
            }
            if (dept.Contains("PE1") && trbl.FinalStatusID == 8) // W-PE-conf
            {
                if (trbl.TroubleFrom.Contains("IN"))
                {
                    trbl.FinalStatusID = 11; //Rejected
                    rejectTo = "DMT";
                }
                else
                {
                    trbl.FinalStatusID = 3; // W-PAE-Chẹck 
                    rejectTo = "PAE";
                }
                // xử lí dữ liệu
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                status = true;
            }
            if (dept.Contains("PE1") && trbl.FinalStatusID == 14) // W-PE-App
            {
                // xử lí dữ liệu
                trbl.FinalStatusID = 8; //W-PE1-Check
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                rejectTo = "PE1";
                status = true;
            }
            if ((dept == "DMT" || dept == "MO") && trbl.FinalStatusID == 10) // W-DMT Check
            {
                if (trbl.TroubleFrom.Contains("IN"))
                {
                    trbl.FinalStatusID = 11; //Rejected
                }
                else
                {
                    trbl.FinalStatusID = 3; // W-PAE Check
                }
                // xử lí dữ liệu
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                status = true;
            }
            if ((dept == "DMT" || dept == "MO") && trbl.FinalStatusID == 7) // W-DMT App
            {
                trbl.FinalStatusID = 10; // W-PAE Check
                // xử lí dữ liệu
                trbl.Progress += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " was rejected TPI";
                trbl.MailContent += System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + Session["Name"] + " was rejected TPI. Reason:" + rejectReason;
                status = true;
            }
            //save


            if (status)
            {
                SaveAndUpdateTPIForm(trbl, null, "", rejectReason, true, false);
                convertToPDF(trbl.Report, trbl.TroubleID);

                if (trbl.FinalStatusID == 11) // reject ve supplier/DMT
                {

                    trbl.IsNeedPURConfirm = false;
                    if (trbl.TroubleFrom.Contains("IN"))
                    {
                        sendEmailJob.sendEmailTrouble(trbl, "DMT,MO", "", "TPI was rejected!!!", sender);
                    }
                    else
                    {
                        sendEmailJob.sendEmailTrouble(trbl, rejectTo, "Check", "TPI was rejected!!!", sender);
                        sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);
                    }
                }

                db.Entry(trbl).State = EntityState.Modified;
                db.SaveChanges();

            }
            else
            {
                if (msg == "")
                {
                    msg = "You can not reject with TPI status: " + trbl.FinalStatusCalogory.FinalStatus + ". Pleas inform to Admin";
                }
            }
            var output = new
            {
                status = status,
                msg = msg
            };

            return Json(output, JsonRequestBehavior.AllowGet);

        }

        public JsonResult reupTPI(int troubleID, HttpPostedFileBase reportFile)
        {
            if (Session["UserID"] == null)
            {
                return Json(false, JsonRequestBehavior.AllowGet);
            }
            var err = "";
            bool status = false;
            var trouble = db.Troubles.Find(troubleID);
            if (reportFile == null)
            {
                err = "Please upload file!";
                goto Exit;
            }
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString();

            var sender = Session["Mail"].ToString();
            var userName = Session["Name"].ToString();
            bool isReject = false;
            if (ModelState.IsValid)
            {
                // Verify TPI form
                // kết quả isPass = true => xử lí tiếp
                // kết quả isPass = false => Trả về  view (nếu do RPA => reject)
                verify verifyTPIForm = verifyTPIFormat(reportFile, trouble.DieNo);
                if (verifyTPIForm.isPass == true)
                {
                    // Kiem tra neu issuer la DMT/MO
                    if (dept.Contains("DMT") || dept.Contains("MO"))
                    {
                        if (trouble.DetailPhenomenon != null && trouble.DetailAction != null && trouble.TroubleTypeID != null && trouble.TroubleAreaID != null
                                    && trouble.FeedbackForR_D != null && trouble.FeedbackForDieSpec != null && trouble.TempCMID != null
                                    && trouble.PerCMID != null && trouble.RootCauseID != null && trouble.NeedDiePO_X6_X7ID != null
                                    && trouble.RenewDecisionID != null && trouble.NeedPEConfirm != null)
                        {
                            // Ko có trường nào null
                            goto continuse;

                        }
                        else
                        {
                            err = "Hãy nhập đủ thông tin ở bảng Sumarize trouble!";
                            goto Exit;
                        }
                    }

                }
                else
                {
                    if (userName.Contains("RPA")) // do RPA upload lên
                    {
                        // Reject
                        isReject = true;
                        err = verifyTPIForm.msg;
                        goto continuse;
                    }
                    else
                    {
                        err = verifyTPIForm.msg;
                        goto Exit;
                    }
                }

                err = verifyTPIForm.msg;
            }

        continuse:
            // 1. Doc file va lay thong tin tu TPI
            trouble = readTPI(trouble, reportFile);


            if (isReject == true)
            {
                trouble.FinalStatusID = 11; // Reject
                trouble.TroubleFrom = "OUTSIDE";

            }
            else
            {
                // 2. Thay doi status
                if (dept == "DMT" || dept == "MO")
                {
                    trouble.TroubleFrom = "INHOUSE";
                    trouble.FinalStatusID = 10; // W-DMT-Check
                }
                else
                {
                    if (db.Die1.Find(trouble.DieID).Belong.Contains("CRG"))
                    {
                        trouble.TroubleFrom = "CRG";
                        trouble.FinalStatusID = 12; // W-CRG-Confirm
                    }
                    else
                    {
                        trouble.TroubleFrom = "OUTSIDE";
                        trouble.FinalStatusID = 3; // W-PAE-Confirm
                    }
                }
            }

            trouble.SubmitDate = DateTime.Now;
            trouble.SubmitBy = Session["Name"].ToString();
            trouble.Progress = today.ToString("yyyy-MM-dd") + ": " + dept + " send report.";
            trouble.Active = true;
            trouble = SaveAndUpdateTPIForm(trouble, reportFile, "Submit", err, isReject, false);
            db.Entry(trouble).State = EntityState.Modified;
            db.SaveChanges();
            status = true;
            convertToPDF(trouble.ReportFromPur, trouble.TroubleID);
            return Json(new { status = status, msg = err }, JsonRequestBehavior.AllowGet);
        //***************************************

        Exit:
            return Json(new { status = status, msg = err }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult checkTPI(int troubleID, string phenomenon, string detailAction, string[] troublType, string[] troubleArea, string[] rootCause,
                                    string[] tempCM, string[] perCM, string poCode, string fbRD, string fbDieSP, string RNdecisionID, string needDMTConfirm,
                                    string needPE1Confirm, string PE1Comment, string PAEComment, string CRGComment, string DMTCommnet, string btn_action,
                                    string dealLine, string tempoAction, string reviseContent, HttpPostedFileBase reportFile)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Trouble_Role"].ToString();
            var trbl = db.Troubles.Find(troubleID);
            var today = DateTime.Now;
            var status = false;
            var err = "";
            var sender = Session["Mail"].ToString();
            if ((btn_action == "PAECheck" || btn_action == "PAEApp" || btn_action == "Revise" || btn_action == "DMTCheck" || btn_action == "DMTApp" || btn_action == "CRGCheck" || btn_action == "CRGApp") && (dept.Contains("PAE") || dept.Contains("DMT") || dept.Contains("MO") || dept.Contains("CRG")))
            {
                if (!String.IsNullOrWhiteSpace(phenomenon) && !String.IsNullOrWhiteSpace(detailAction) && troublType?.Length > 0 && troubleArea?.Length > 0
                            && rootCause?.Length > 0 && tempCM?.Length > 0 && perCM?.Length > 0
                            && !String.IsNullOrWhiteSpace(poCode) && !String.IsNullOrWhiteSpace(fbRD) && !String.IsNullOrWhiteSpace(fbDieSP)
                            && !String.IsNullOrWhiteSpace(RNdecisionID) && !String.IsNullOrWhiteSpace(needDMTConfirm) && !String.IsNullOrWhiteSpace(needPE1Confirm))
                {
                    if (trbl.TroubleFrom.Contains("OUT"))
                    {
                        if (trbl.Report == null && reportFile == null)
                        {
                            err = "Please update file TPI/Report!";
                            goto exitLoop;
                        }
                    }
                    // Xu li du lieu
                    trbl.DetailPhenomenon = phenomenon;
                    trbl.DetailAction = detailAction;
                    trbl.TroubleTypeID = String.Join(",", troublType);
                    trbl.TroubleAreaID = String.Join(",", troubleArea);
                    trbl.FeedbackForR_D = fbRD == "true" ? true : false;
                    trbl.FeedbackForDieSpec = fbDieSP == "true" ? true : false;
                    trbl.TempCMID = String.Join(",", tempCM);
                    trbl.PerCMID = String.Join(",", perCM);
                    trbl.RootCauseID = String.Join(",", rootCause);
                    trbl.NeedDiePO_X6_X7ID = int.Parse(poCode);
                    trbl.RenewDecisionID = int.Parse(RNdecisionID);
                    trbl.NeedPEConfirm = needPE1Confirm == "true" ? true : false;
                    trbl.NeedDMTConfirm = needDMTConfirm == "true" ? true : false;
                    //**********************************************************************************************************************************
                    //PAE Confirm
                    if (btn_action == "PAECheck" && dept.Contains("PAE") && trbl.FinalStatusID == 3) // W-PAE-Conf
                    {
                        trbl.PIC = Session["Name"].ToString();
                        trbl.FeedbackDate = today;
                        trbl.FinalStatusID = 9; // W-PAE-App
                        trbl.PAEComment = PAEComment;
                        // File report handler ==> start
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);


                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }

                        }
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PAE checked";
                        // update & convert to PDF
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "PAECheck", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);
                        goto DoneLable;
                    }
                    //**********************************************************************************************************************************
                    //**********************************************************************************************************************************
                    //PAE App
                    if (btn_action == "PAEApp" && role == "Approve" && trbl.FinalStatusID == 9) // W-PAE-App
                    {
                        trbl.PAECommentBy = Session["Name"].ToString();
                        trbl.PAECommentDate = today;
                        trbl.PAEComment = PAEComment;

                        // File report handler ==> start
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);


                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                        }
                        // update & convert to PDF
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "PAEApp", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);
                        //File report handler ==> end
                        // Phân loại Route
                        //**********************************************************************
                        trbl = StatusChangeAfterPEorPAEAppOrCRGApp(trbl);
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PAE Approved";
                        //********************************************************************
                        goto DoneLable;
                    }
                    //***********************************************************************************************************************************
                    //DMT/QE/Check
                    if (btn_action == "DMTCheck" && (dept.Contains("DMT") || dept.Contains("MO")) && trbl.FinalStatusID == 10) // W-DMT/QE Check
                    {

                        trbl.DMTCheckerComment = DMTCommnet;
                        trbl.DMTCheckBy = Session["Name"].ToString();
                        trbl.DMTCheckDate = today;
                        trbl.FinalStatusID = 7; // W-DMT G6 App
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": QE/DMT Checked";
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);
                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                        }
                        // update & convert to PDF
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "DMTCheck", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);

                        goto DoneLable;
                    }
                    //***************************************************************************************************************************************
                    //***********************************************************************************************************************************
                    // DMT APP
                    if (btn_action == "DMTApp" && role == "Approve" && dept.Contains("DMT") && trbl.FinalStatusID == 7) //W-DMT-App
                    {

                        trbl.DMTComment = DMTCommnet;
                        trbl.DMTAppBy = Session["Name"].ToString();
                        trbl.DMTAppDate = today;
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": DMT Approved";
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);
                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                        }
                        // update & convert to PDF
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "DMTApp", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);
                        trbl = StatusChangeAfterPEorPAEAppOrCRGApp(trbl);
                        goto DoneLable;
                    }
                    //***************************************************************************************************************************************
                    //***************************************************************************************************************************************
                    if (btn_action == "CRGCheck" && dept.Contains("CRG") && trbl.FinalStatusID == 12) // W-CRG-Conf
                    {
                        trbl.CRGComment = CRGComment;
                        trbl.CRGCheckBy = Session["Name"].ToString();
                        trbl.CRGCheckDate = today;
                        trbl.FinalStatusID = 13; // W-CRG-App

                        // File report handler ==> start
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);
                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                            // update & convert to PDF
                        }
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "CRGCheck", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);

                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": CRG checked";
                        goto DoneLable;
                        //File report handler ==> end

                    }
                    //****************************************************************************************************************************************
                    //****************************************************************************************************************************************
                    //PAE App
                    if (btn_action == "CRGApp" && role == "Approve" && trbl.FinalStatusID == 13) // W-CRG-App
                    {
                        trbl.CRGAppBy = Session["Name"].ToString();
                        trbl.CRGAppDate = today;
                        trbl.CRGComment = CRGComment;
                        // File report handler ==> start
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);
                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                            // update & convert to PDF
                        }
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "CRGApp", "", false, false);
                        convertToPDF(trbl.Report, trbl.TroubleID);
                        //File report handler ==> end
                        // Phân loại Route
                        //**********************************************************************
                        trbl = StatusChangeAfterPEorPAEAppOrCRGApp(trbl);
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": CRG Approved";
                        //********************************************************************
                        goto DoneLable;

                    }

                    //***************************************************************************************************************************************
                    //***************************************************************************************************************************************
                    if (btn_action == "Revise" && (dept.Contains("PAE") || dept.Contains("DMT") || dept.Contains("MO") || dept.Contains("CRG")))
                    {
                        // Upgrage for TroubleNo issue No *********************
                        // PUR - 00 => PAE 1st 01 => revise 02,03...
                        trbl.TroubleNo = genarateTPINo(trbl, false);
                        if (dept.Contains("PAE"))
                        {
                            trbl.PAEComment = PAEComment;
                        }
                        if (dept.Contains("DMT") || dept.Contains("MO"))
                        {
                            trbl.DMTComment = DMTCommnet;
                        }
                        if (dept.Contains("CRG"))
                        {
                            trbl.CRGComment = CRGComment;
                        }

                        // File report handler ==> start
                        if (reportFile != null)
                        {
                            verify verifyTPIForm = verifyTPIFormat(reportFile, trbl.DieNo);
                            if (verifyTPIForm.isPass == false)
                            {
                                err = verifyTPIForm.msg;
                                goto exitLoop;
                            }
                        }
                        else
                        {
                            //File file Cũ => đổi tên theo số trouble Mới
                            string pathFileOld = Server.MapPath("~/File/TroubleReport/" + trbl.Report);
                            string fileExt = Path.GetExtension(pathFileOld);
                            string newFileName = "[" + trbl.TroubleNo + "] " + trbl.DieNo + fileExt;
                            trbl.Report = newFileName;
                            System.IO.File.Copy(pathFileOld, Server.MapPath("~/File/TroubleReport/" + newFileName), true);
                        }
                        // update & convert to PDF
                        trbl = SaveAndUpdateTPIForm(trbl, reportFile, "", reviseContent, false, true);
                        convertToPDF(trbl.Report, trbl.TroubleID);
                        trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": " + dept + " Revised";
                        trbl.ReviseDate = today;
                        trbl.ReviseBy = Session["Name"].ToString();

                        if (dept == "DMT" || dept == "MO")
                        {
                            trbl.FinalStatusID = 7; // W-DMT G6 Apptrbl.FinalStatusID == 
                            trbl.DMTCheckBy = Session["Name"].ToString() + " [Revised]";
                            trbl.DMTCheckDate = today;
                        }
                        else
                        {
                            if (trbl.Die1.Belong == "CRG")
                            {
                                trbl.FinalStatusID = 13; // W-PAE-App
                                trbl.CRGCheckBy = Session["Name"].ToString() + " [Revised]";
                                trbl.CRGCheckDate = today;
                            }
                            else
                            {
                                trbl.FinalStatusID = 9; // W-PAE-App
                                trbl.PIC = Session["Name"].ToString() + " [Revised]";
                                trbl.FeedbackDate = today;
                            }
                        }
                    }


                }
                else
                {
                    status = false;
                    err = "Please input all information!";
                    goto exitLoop;
                }
            }
            if (btn_action == "PE1Check" && dept.Contains("PE1") && trbl.FinalStatusID == 8)
            {
                trbl.PE1Comment = PE1Comment;
                trbl.PE1CommentBy = Session["Name"].ToString();
                trbl.PE1CommentDate = today;
                trbl.FinalStatusID = 14; // W-PE1-App

                trbl = SaveAndUpdateTPIForm(trbl, null, "PE1Check", "", false, false);
                convertToPDF(trbl.Report, trbl.TroubleID);

                trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PE1 checked";
                goto DoneLable;
                //File report handler ==> end
            }

            if (btn_action == "PE1App" && dept.Contains("PE1"))
            {
                trbl.PE1Comment = PE1Comment;
                trbl.PE1AppBy = Session["Name"].ToString();
                trbl.PE1AppDate = today;
                trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PE1 Approved";
                trbl = SaveAndUpdateTPIForm(trbl, null, "PE1App", "", false, false);
                convertToPDF(trbl.Report, trbl.TroubleID);
                // Phân loại Route
                //**********************************************************************
                trbl = StatusChangeAfterPEorPAEAppOrCRGApp(trbl);
                goto DoneLable;
                //File report handler ==> end
            }

            if (btn_action == "PURCheck" && dept.Contains("PUR"))
            {
                trbl.IsNeedPURConfirm = false;

                trbl.TemapratyAction = !String.IsNullOrWhiteSpace(tempoAction) ? tempoAction : trbl.TemapratyAction;
                trbl.DeadLineIfUrgent = !String.IsNullOrWhiteSpace(dealLine) ? DateTime.Parse(dealLine) : trbl.DeadLineIfUrgent;
                trbl.PURCommentBy = Session["Name"].ToString();
                trbl.PURCommentDate = today;
                trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PUR Confirmed";
                trbl = SaveAndUpdateTPIForm(trbl, null, "PURCheck", "", false, false);
                convertToPDF(trbl.Report, trbl.TroubleID);
                // Phân loại Route
                //**********************************************************************
                if (trbl.FinalStatusID == 1 || trbl.FinalStatusID == 2 || trbl.FinalStatusID == 4 || trbl.FinalStatusID == 5 || trbl.FinalStatusID == 6 || trbl.FinalStatusID == 11 || trbl.FinalStatusID == 15)
                {
                    sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);
                }


                goto DoneLable;
            }

            if (btn_action == "Revise" && (dept.Contains("PE1") || dept.Contains("PUR")))
            {
                if (dept.Contains("PE1"))
                {
                    trbl.PE1Comment = reviseContent;
                    trbl.PE1CommentBy = Session["Name"].ToString();
                    trbl.PE1CommentDate = today;
                    trbl.FinalStatusID = 14; // W-PE1-App
                    trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PE1 Revised";
                }
                if (dept.Contains("PUR"))
                {
                    trbl.TemapratyAction = !String.IsNullOrWhiteSpace(tempoAction) ? tempoAction : trbl.TemapratyAction;
                    trbl.DeadLineIfUrgent = !String.IsNullOrWhiteSpace(dealLine) ? DateTime.Parse(dealLine) : trbl.DeadLineIfUrgent;
                    trbl.PURCommentBy = Session["Name"].ToString();
                    trbl.PURCommentDate = today;
                    trbl.Progress = trbl.Progress + System.Environment.NewLine + today.ToString("yyyy-MM-dd") + ": PUR Revised";
                }

                // Upgrage for TroubleNo issue No *********************
                // PUR - 00 => PAE 1st 01 => revise 02,03...
                trbl.TroubleNo = genarateTPINo(trbl, false);
                // File report handler ==> start
                //File file Cũ => đổi tên theo số trouble Mới
                string pathFileOld = Server.MapPath("~/File/TroubleReport/" + trbl.Report);
                string fileExt = Path.GetExtension(pathFileOld);
                string newFileName = "[" + trbl.TroubleNo + "] " + trbl.DieNo + fileExt;
                trbl.Report = newFileName;
                System.IO.File.Copy(pathFileOld, Server.MapPath("~/File/TroubleReport/" + newFileName), true);

                // update & convert to PDF
                trbl = SaveAndUpdateTPIForm(trbl, null, "", reviseContent, false, true);
                convertToPDF(trbl.Report, trbl.TroubleID);
                trbl.ReviseDate = today;
                trbl.ReviseBy = Session["Name"].ToString();
                goto DoneLable;
            }

        DoneLable:
            // save
            db.Entry(trbl).State = EntityState.Modified;
            db.SaveChanges();
            status = true;
            //send email
            if (trbl.FinalStatusID == 7) // W-DMT-App
            {
                sendEmailJob.sendEmailTrouble(trbl, "DMT", "Approve", "", sender);
            }
            if (trbl.FinalStatusID == 3 || trbl.FinalStatusID == 9) // W-PAE-Check
            {
                sendEmailJob.sendEmailTrouble(trbl, "PAE", "", "", sender);
            }

            if (trbl.FinalStatusID == 8) // W-PE-App
            {
                sendEmailJob.sendEmailTrouble(trbl, "PE1", "", "", sender);
            }
            if (trbl.FinalStatusID == 4 || trbl.FinalStatusID == 6) // W-MR || FA result
            {
                if (trbl.TroubleFrom.Contains("IN"))
                {
                    sendEmailJob.sendEmailTrouble(trbl, "DMT,MO", "", "Please start Repair/OH/Spare", sender);
                }
                else
                {
                    sendEmailJob.sendEmailTrouble(trbl, "PUR", "", "Please information to supplier TPI result", sender);
                }
            }
            if (trbl.FinalStatusID == 1) // Close
            {
                sendEmailJob.sendEmailTrouble(trbl, "PUR", "", "No repair this trouble", sender);
            }
            if (trbl.FinalStatusID == 13) // W-CRG-App
            {
                sendEmailJob.sendEmailTrouble(trbl, "CRG", "Approve", "", sender);
            }


            return Json(new { status = status, msg = err }, JsonRequestBehavior.AllowGet);
        exitLoop:
            return Json(new { status = status, msg = err }, JsonRequestBehavior.AllowGet);
        }

        public class verify
        {
            public bool isPass { set; get; }
            public string msg { set; get; }
        }
        public verify verifyTPIFormat(HttpPostedFileBase reportFile, string dieNoCheck)
        {
            // Trả về kết quả verify
            // Nhiệm vụ của verify 
            // 1. Check format đúng version hay ko?
            // 2. check thông tin đã nhập đầy đủ hay ko?
            // 3. Check Die tồn tại hay ko?

            bool isPass = false;
            string msg = "";
            if (reportFile == null)
            {
                isPass = false;
                msg = "No File upload";
                goto Exit;
            }



            // 1. Check format đúng version chưa
            var Y1 = "LPE-0010";
            var Y2 = "Att.06/Rev 01";
            var R11 = "Final Decision";
            var G17 = "PAE confirm";
            var B59 = "For Inhouse/CRG/PUR";
            var U63 = "Information refer";


            using (ExcelPackage package = new ExcelPackage(reportFile.InputStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                // Check version

                var y1 = worksheet.Cells["Y1"].Text.Trim();
                var y2 = worksheet.Cells["Y2"].Text.Trim();

                //Check hàng

                var r11 = worksheet.Cells["R11"].Text.Trim();
                var g17 = worksheet.Cells["G17"].Text.Trim();
                var b59 = worksheet.Cells["B59"].Text.Trim();
                var u63 = worksheet.Cells["U63"].Text.Trim();

                if ((Y1 != y1 || Y2 != y2))
                {
                    msg = "Wrong Format, Collunm Y1 = " + Y1 + " & Collunm Y2 = " + Y2;
                    isPass = false;
                    goto Exit;
                }
                else
                {
                    if ((R11 != r11 || G17 != g17 || B59 != b59 || U63 != u63))
                    {
                        msg = "Wrong Form, Format was changed No of Collunm or No of Row";
                        isPass = false;
                        goto Exit;
                    }
                    else
                    {
                        // Format OK
                        // 2.Check Đã đủ thông tin chưa
                        var SubmitType = worksheet.Cells["F4"].Text.Trim();
                        var classification = worksheet.Cells["F5"].Text.Trim();
                        var dieNo = worksheet.Cells["F6"].Text.Trim();
                        var partNo = worksheet.Cells["F7"].Text.Trim();
                        var cavNG = worksheet.Cells["F8"].Text.Trim();
                        var TotalCav = worksheet.Cells["F9"].Text.Trim();
                        var MCsize = worksheet.Cells["F10"].Text.Trim();
                        var ecn_eriNo = worksheet.Cells["P4"].Text.Trim();
                        var happenDate = worksheet.Cells["P5"].Text.Trim();
                        var shot = worksheet.Cells["P6"].Text.Trim();
                        var prodStatus = worksheet.Cells["P7"].Text.Trim();
                        var phase = worksheet.Cells["P8"].Text.Trim();
                        var proccesNG = worksheet.Cells["P9"].Text.Trim();
                        var dieMaterial = worksheet.Cells["P10"].Text.Trim();
                        var supplierCode = worksheet.Cells["W3"].Text.Trim();
                        var TPINo = worksheet.Cells["L3"].Text.Trim();
                        var dueDate = worksheet.Cells["X10"].Text.Trim();

                        if (String.IsNullOrWhiteSpace(SubmitType) || String.IsNullOrWhiteSpace(classification) || String.IsNullOrWhiteSpace(dieNo) ||
                            String.IsNullOrWhiteSpace(partNo) || String.IsNullOrWhiteSpace(cavNG) || String.IsNullOrWhiteSpace(TotalCav) ||
                            String.IsNullOrWhiteSpace(MCsize) || String.IsNullOrWhiteSpace(shot) || String.IsNullOrWhiteSpace(prodStatus) ||
                            String.IsNullOrWhiteSpace(phase) || String.IsNullOrWhiteSpace(proccesNG) || String.IsNullOrWhiteSpace(dieMaterial) || String.IsNullOrWhiteSpace(supplierCode))
                        {
                            msg = "Please input all field (*)";
                            isPass = false;
                            goto Exit;
                        }

                        var existDie = db.Die1.Where(x => x.DieNo == dieNo && x.Active != false && x.isCancel != true).FirstOrDefault();
                        if (existDie == null)
                        {
                            msg = "Die ID: " + dieNo + " not correct.";
                            isPass = false;
                            goto Exit;
                        }


                        if (!String.IsNullOrEmpty(dieNoCheck))
                        {
                            if (dieNoCheck != dieNo)
                            {
                                msg = "Maybe you select wrong TPI, Die ID on TPI: " + dieNo + " but you'r checking TPI for Die ID " + dieNoCheck;
                                isPass = false;
                                goto Exit;
                            }
                        }


                        if (SubmitType == "Revise/Feedback")
                        {
                            // Check Old TPI đúng dieNo ko?
                            var oldTPI = db.Troubles.Where(x => x.TroubleNo == TPINo).FirstOrDefault();
                            if (oldTPI != null)
                            {
                                if (oldTPI.DieNo != dieNo)
                                {
                                    msg = SubmitType + " for TPI_No: " + TPINo + "But DieID input not correct!";
                                    isPass = false;
                                    goto Exit;
                                }
                                else
                                {
                                    msg = "OK";
                                    isPass = true;
                                }
                            }
                            else
                            {
                                msg = "Submit Type* = " + SubmitType + " but TPI_No Wrong or Not Input, Plz check again!";
                                isPass = false;
                                goto Exit;
                            }
                        }
                        else
                        {
                            if (SubmitType == "New_Submit")
                            {
                                var isExistDie = db.Die1.Where(x => x.DieNo == dieNo && x.Disposal != true && x.Active != false).FirstOrDefault();
                                if (isExistDie != null)
                                {
                                    msg = "OK";
                                    isPass = true;
                                }
                                else
                                {
                                    msg = "Please re-check DieID input correct or not? System not contain this Die ID";
                                    isPass = true;
                                    goto Exit;
                                }
                            }

                        }

                        if (classification == "ECN/ERI")
                        {
                            if (String.IsNullOrWhiteSpace(ecn_eriNo))
                            {
                                msg = "Classification* = " + classification + " but you not input ENC/ERINo. Plz input it!";
                                isPass = false;
                                goto Exit;
                            }
                        }

                    }
                }



            }

        Exit:
            var output = new verify()
            {
                isPass = isPass,
                msg = msg
            };
            return output;
        }

        public Trouble readTPI(Trouble trouble, HttpPostedFileBase reportFile)
        {
            if (reportFile == null)
            {
                return trouble;
            }
            var today = DateTime.Now;
            using (ExcelPackage package = new ExcelPackage(reportFile.InputStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                // 2.Check Đã đủ thông tin chưa
                var SubmitType = worksheet.Cells["F4"].Text.Trim();
                var classification = worksheet.Cells["F5"].Text.Trim();
                var dieNo = worksheet.Cells["F6"].Text.Trim();
                var partNo = worksheet.Cells["F7"].Text.Trim();
                var cavNG = worksheet.Cells["F8"].Text.Trim();
                var TotalCav = worksheet.Cells["F9"].Text.Trim();
                var MCsize = worksheet.Cells["F10"].Text.Trim();
                var ecn_eriNo = worksheet.Cells["P4"].Text.Trim();
                var happenDate = worksheet.Cells["P5"].Text.Trim();
                var shot = worksheet.Cells["P6"].Text.Trim();
                var prodStatus = worksheet.Cells["P7"].Text.Trim();
                var phase = worksheet.Cells["P8"].Text.Trim();
                var proccesNG = worksheet.Cells["P9"].Text.Trim();
                var dieMaterial = worksheet.Cells["P10"].Text.Trim();
                var supplierCode = worksheet.Cells["W3"].Text.Trim();
                var TPINo = worksheet.Cells["L3"].Text.Trim();
                var dueDate = worksheet.Cells["X10"].Text.Trim();



                var existDie = db.Die1.Where(x => x.DieNo == dieNo && x.Active != false).FirstOrDefault();

                if (SubmitType.Contains("Revise/Feedback") && !String.IsNullOrWhiteSpace(TPINo))
                {
                    var existTrouble = db.Troubles.Where(x => x.TroubleNo == TPINo && x.Active != false).FirstOrDefault();
                    if (existTrouble != null)
                    {
                        trouble = existTrouble;
                    }
                }

                if (existDie == null)
                {

                    trouble.DieID = 14389991; // Die fake Offical DB
                    //trouble.DieID = 14389940; // Die fake local DB for testing
                    trouble.Die1 = db.Die1.Find(14389940);
                }
                else
                {
                    trouble.DieID = existDie.DieID;
                }

                trouble.PartNo = partNo.ToUpper();
                trouble.SubmitType = SubmitType;
                trouble.TroubleName = classification;
                trouble.DieNo = dieNo;
                try
                {
                    trouble.HappenDate = DateTime.Parse(happenDate);
                }
                catch
                {
                    trouble.HappenDate = today;
                }

                try
                {
                    trouble.ActualShort = double.Parse(shot);
                }
                catch
                {
                    trouble.ActualShort = 0;
                }


                trouble.DieTroubleRunningStatusID = prodStatus.Contains("Stop") ? 1 : prodStatus.Contains("Rework") ? 2 : prodStatus.Contains("Sorting") ? 3 : prodStatus.Contains("Temporary") ? 4 : prodStatus.Contains("other die") ? 5 : 6;
                trouble.Phase = phase;
                try
                {
                    trouble.SupplierName = db.Suppliers.Where(x => x.SupplierCode == supplierCode).FirstOrDefault().SupplierName;
                }
                catch
                {
                    trouble.SupplierName = existDie?.Supplier.SupplierName;
                }

                try
                {
                    trouble.DeadLineIfUrgent = DateTime.Parse(dueDate);
                }
                catch
                {
                    //
                }
                try
                {
                    if (trouble.TroubleLevelID == null && trouble.DeadLineIfUrgent != null)
                    {
                        trouble.TroubleLevelID = trouble.DeadLineIfUrgent.Value.Subtract(today).Days > 14 ? 1 : trouble.DeadLineIfUrgent.Value.Subtract(today).Days > 7 ? 2 : 3;
                    }
                    else
                    {
                        trouble.TroubleLevelID = 1;
                    }
                }
                catch
                {
                    //
                }

                // Update need PUR confirm hay ko?
                //if (!trouble.TroubleFrom.Contains("IN") && trouble.Phase.Contains("MP") && (trouble.TroubleName.Contains("ECN/ERI") || trouble.TroubleName.Contains("FA improvement") || trouble.TroubleName.Contains("Die transfer") || trouble.TroubleName.Contains("MP trouble")))
                //{
                //    trouble.IsNeedPURConfirm = true;
                //    if (trouble.PURCommentDate != null)
                //    {
                //        trouble.IsNeedPURConfirm = false;
                //    }
                //}
                //else
                //{
                //    trouble.IsNeedPURConfirm = false;
                //}


            }

            return trouble;
        }

        public string genarateTPINo(Trouble trouble, bool isNewSubmit)
        {
            var TotalTrbCase = db.Troubles.Count() + 1;
            var today = DateTime.Now;
            var TPINo = "";
            // Tu 2024 so TPI se duoc danh so lai theo nam
            if (today > new DateTime(2023, 12, 31))
            {
                TotalTrbCase = db.Troubles.Where(x => x.SubmitDate.Value.Year == today.Year).Count() + 1;
            }

            if (isNewSubmit == true)
            {
                TPINo = "TPI" + today.ToString("yyMMdd") + "-" + TotalTrbCase + "-00";
            }
            else
            {
                try
                {
                    if (trouble.TroubleNo != null)
                    {
                        // PUR - 00 => PAE 1st 01 => revise 02,03...
                        var CurrentTroubleNo = trouble.TroubleNo;
                        var upver = trouble.TroubleNo.Substring(CurrentTroubleNo.Length - 2, 2);
                        int upverInt = Convert.ToInt16(upver) + 1;
                        string upverStr = Convert.ToString(upverInt);
                        if (upverStr.Length == 1)
                        {
                            upverStr = "0" + upverStr;
                        }
                        var mainNo = CurrentTroubleNo.Remove(CurrentTroubleNo.Length - 2, 2);
                        TPINo = mainNo + upverStr;
                    }
                }
                catch
                {

                }

            }
            return TPINo;
        }

        public Trouble SaveAndUpdateTPIForm(Trouble trouble, HttpPostedFileBase reportFile, string progress, string msg, bool isReject, bool isRevise)
        {
            var today = DateTime.Now;
            var fileName = "";
            var dept = Session["Dept"].ToString();
            trouble = readTPI(trouble, reportFile);
            if (reportFile != null)
            {
                //1. Luu file vat li
                fileName = "[" + trouble.TroubleNo + "] " + trouble.DieNo;
                string fileExt = Path.GetExtension(reportFile.FileName);
                string path = Server.MapPath("~/File/TroubleReport/");
                fileName += fileExt;
                reportFile.SaveAs(path + Path.GetFileName(fileName));
                // removePassword(path + Path.GetFileName(fileName));
            }
            else
            {
                fileName = !String.IsNullOrWhiteSpace(trouble.Report) ? trouble.Report : trouble.ReportFromPur;
            }

            //2. Updat file Vat li
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/TroubleReport/" + fileName))))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                var partNo = worksheet.Cells["F7"].Text.Trim().ToUpper();
                worksheet.Cells["L3"].Value = trouble.TroubleNo;
                worksheet.Cells["W6"].Value = db.CommonDie1.Where(x => x.PartNo == partNo && x.Active != false && x.Die1.Disposal != true && x.Die1.Active != false).Count();
                worksheet.Cells["W7"].Value = db.Parts1.Where(x => x.PartNo == partNo).FirstOrDefault()?.Model;
                worksheet.Cells["I14"].Value = "";
                if (isReject) // Reject
                {
                    worksheet.Cells["M13"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]Rejected: " + msg + System.Environment.NewLine + worksheet.Cells["M13"].Text;
                    worksheet.Cells["I14"].Value = "REJECTED";

                    trouble.Report = fileName;
                    if (Session["Name"].ToString() == "RPA")
                    {
                        worksheet.Cells["S51"].Value = "";
                    }
                }
                if (isRevise) // Revise
                {
                    if (dept.Contains("PAE") || dept.Contains("CRG") || dept.Contains("DMT") || dept.Contains("MO"))
                    {
                        worksheet.Cells["M13"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]Revised: " + msg + System.Environment.NewLine + worksheet.Cells["M13"].Text;
                        if (dept.Contains("PAE"))
                        {
                            trouble.Report = fileName;
                            worksheet.Cells["N60"].Value = Session["Name"].ToString() + System.Environment.NewLine + "[Revised]";
                            worksheet.Cells["N62"].Value = today.ToString("yyyy-MM-dd");

                        }
                        else
                        {
                            trouble.Report = fileName;
                            worksheet.Cells["F61"].Value = Session["Name"].ToString() + System.Environment.NewLine + "[Revised]";
                            worksheet.Cells["F62"].Value = today.ToString("yyyy-MM-dd");

                        }
                    }
                    else
                    {
                        worksheet.Cells["M16"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]" + "Revised: " + msg + System.Environment.NewLine + worksheet.Cells["M16"].Text;

                    }
                    worksheet.Cells["I14"].Value = "REVISE";

                }
                if (progress.Contains("Submit"))
                {
                    trouble.ReportFromPur = fileName;
                    worksheet.Cells["B61"].Value = Session["Name"].ToString();
                    worksheet.Cells["B62"].Value = today.ToString("yyyy-MM-dd");
                    worksheet.Cells["U11"].Value = "";
                }
                if (progress.Contains("PAECheck"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["N60"].Value = Session["Name"].ToString();
                    worksheet.Cells["N62"].Value = today.ToString("yyyy-MM-dd");


                }
                if (progress.Contains("PAEApp"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["Q60"].Value = Session["Name"].ToString();
                    worksheet.Cells["Q62"].Value = today.ToString("yyyy-MM-dd");

                    if (!String.IsNullOrWhiteSpace(trouble.PAEComment))
                    {
                        worksheet.Cells["M16"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]PAE: " + trouble.PAEComment + System.Environment.NewLine + worksheet.Cells["M16"].Text;
                    }
                }
                if (progress.Contains("PE1Check"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["T60"].Value = Session["Name"].ToString();
                    worksheet.Cells["T62"].Value = today.ToString("yyyy-MM-dd");

                }
                if (progress.Contains("PE1App"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["W60"].Value = Session["Name"].ToString();
                    worksheet.Cells["W62"].Value = today.ToString("yyyy-MM-dd");
                    if (!String.IsNullOrWhiteSpace(trouble.PE1Comment))
                    {
                        worksheet.Cells["M16"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]PE1: " + trouble.PE1Comment + System.Environment.NewLine + worksheet.Cells["M16"].Text;
                    }
                }
                if (progress.Contains("DMTCheck") || progress.Contains("CRGCheck"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["F61"].Value = Session["Name"].ToString();
                    worksheet.Cells["F62"].Value = today.ToString("yyyy-MM-dd");

                }
                if (progress.Contains("DMTApp") || progress.Contains("CRGApp"))
                {
                    trouble.Report = fileName;
                    worksheet.Cells["J61"].Value = Session["Name"].ToString();
                    worksheet.Cells["J62"].Value = today.ToString("yyyy-MM-dd");

                    if (!String.IsNullOrWhiteSpace(trouble.DMTComment) && progress.Contains("DMTApp"))
                    {
                        worksheet.Cells["M13"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]DMT: " + trouble.DMTComment + System.Environment.NewLine + worksheet.Cells["M13"].Text;
                    }
                    if (!String.IsNullOrWhiteSpace(trouble.CRGComment) && progress.Contains("CRGApp"))
                    {
                        worksheet.Cells["M16"].Value = "[" + today.ToString("yyyy/MM/dd") + Session["Name"].ToString() + "]CRG: " + trouble.CRGComment + System.Environment.NewLine + worksheet.Cells["M16"].Text;
                    }
                }
                if (progress.Contains("PURCheck"))
                {
                    trouble.Report = fileName;
                    if (!String.IsNullOrWhiteSpace(trouble.TemapratyAction))
                    {
                        worksheet.Cells["M16"].Value = "[" + today.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + "]PUR: " + trouble.TemapratyAction + System.Environment.NewLine + worksheet.Cells["M16"].Text;
                    }
                }
                worksheet.Cells["U11"].Value = db.NeedDiePOCalogories.Find(trouble.NeedDiePO_X6_X7ID) == null ? "" : db.NeedDiePOCalogories.Find(trouble.NeedDiePO_X6_X7ID).PODieCode;
                worksheet.Cells["J17"].Value = db.RenewDecisionCalogories.Find(trouble.RenewDecisionID) == null ? "" : db.RenewDecisionCalogories.Find(trouble.RenewDecisionID).DecisionStatus;

                package.Save();
            }



            return trouble;

        }


        public Trouble StatusChangeAfterPEorPAEAppOrCRGApp(Trouble trbl)
        {
            var today = DateTime.Now;

            // if stop 
            if (trbl.NeedDiePO_X6_X7ID == 10) // Stop
            {
                trbl.FinalStatusID = 1; // Close

                trbl.CloseDate = today;
                trbl.CloseContent = "Stop using die";
                trbl.isNeedRPA = true;
                // Update die
                var die = db.Die1.Find(trbl.DieID);
                die.DieStatusID = 8; // Stop
                die.RemarkDieStatusUsing = "Die Stop Use follow TPI: " + trbl.TroubleNo;
                die.Short = trbl.ActualShort;
                die.RecordDate = today;
                db.Entry(die).State = EntityState.Modified;
                db.SaveChanges();

                trbl.IsNeedPURConfirm = false;
                sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);
                return trbl;
            }
            // if stop 
            //if (trbl.TroubleName.Contains("Design Change")) // 
            //{
            //    trbl.FinalStatusID = 1; // Close

            //    trbl.CloseDate = today;
            //    trbl.CloseContent = "Stop using die";
            //    trbl.isNeedRPA = true;
            //    return trbl;
            ////}



            // AfterPAEApp
            if (trbl.FinalStatusID == 9)  // current w-pae-app
            {
                if (trbl.NeedDMTConfirm == true)
                {
                    trbl.FinalStatusID = 10; // W-DMT-Check
                }
                else
                {
                    if (trbl.NeedPEConfirm == true)
                    {
                        trbl.FinalStatusID = 8; // W-PE1-Check
                    }
                    else // KO cần xác nhận 
                    {
                        if (trbl.NeedDiePO_X6_X7ID == 1 || trbl.NeedDiePO_X6_X7ID == 2 || trbl.NeedDiePO_X6_X7ID == 7 || trbl.NeedDiePO_X6_X7ID == 8 || trbl.NeedDiePO_X6_X7ID == 9) // X7 or X6, x6OH,X5, X9
                        {
                            if (trbl.NeedDiePO_X6_X7ID == 8) // OH
                            {
                                trbl.FinalStatusID = 1; // Close
                            }
                            else
                            {
                                trbl.FinalStatusID = checkMRPOissue(trbl);
                            }


                        }
                        if (trbl.NeedDiePO_X6_X7ID == 3 || trbl.NeedDiePO_X6_X7ID == 6 || trbl.NeedDiePO_X6_X7ID == 11) // Maker Respone || DMT repair || Use spare
                        {
                            trbl.FinalStatusID = 6; //W-FA-RESULT

                        }

                        if (trbl.NeedDiePO_X6_X7ID == 4 || trbl.NeedDiePO_X6_X7ID == 5) // No Need PO
                        {
                            trbl.FinalStatusID = 1; // Close
                            trbl.CloseDate = today;
                            trbl.CloseContent = "No Repair die";

                        }
                        if (trbl.IsNeedPURConfirm == true)
                        {
                            trbl.isNeedRPA = false;
                        }
                        else
                        {
                            if (trbl.TroubleFrom.Contains("OUT"))
                            {
                                trbl.isNeedRPA = true;
                                sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);
                            }
                            else
                            {
                                trbl.isNeedRPA = false;
                            }

                        }

                    }
                }
                goto Exit;
            }

            // AfterCRGApp

            if (trbl.FinalStatusID == 13)  // current w-CRG-app
            {
                if (trbl.NeedDiePO_X6_X7ID == 1 || trbl.NeedDiePO_X6_X7ID == 2 || trbl.NeedDiePO_X6_X7ID == 7 || trbl.NeedDiePO_X6_X7ID == 8 || trbl.NeedDiePO_X6_X7ID == 9) // X7 or X6, x6OH,X5, X9
                {
                    trbl.FinalStatusID = checkMRPOissue(trbl);

                }
                if (trbl.NeedDiePO_X6_X7ID == 3 || trbl.NeedDiePO_X6_X7ID == 6 || trbl.NeedDiePO_X6_X7ID == 11) // Maker Respone || DMT repair || Use spare
                {
                    trbl.FinalStatusID = 6; //W-FA-RESULT

                }

                if (trbl.NeedDiePO_X6_X7ID == 4 || trbl.NeedDiePO_X6_X7ID == 5) // No Need PO
                {
                    trbl.FinalStatusID = 1; // Close
                    trbl.CloseDate = today;
                    trbl.CloseContent = "No Repair die";

                }
                if (trbl.IsNeedPURConfirm == true)
                {
                    trbl.isNeedRPA = false;
                }
                else
                {
                    trbl.isNeedRPA = true;
                    sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);

                }


                goto Exit;
            }

            // AfterDMTApp
            if (trbl.FinalStatusID == 7) // w-DMT-app
            {
                if (trbl.NeedDiePO_X6_X7ID == 7 || trbl.NeedDiePO_X6_X7ID == 8) // X6_OH & X9 => Need PAE check
                {
                    trbl.FinalStatusID = 9; // W-PAE-App

                }
                else
                {
                    if (trbl.NeedPEConfirm == true)
                    {
                        trbl.FinalStatusID = 8; // W-PE1-App
                    }
                    else
                    {
                        if (trbl.NeedDiePO_X6_X7ID == 1 || trbl.NeedDiePO_X6_X7ID == 2 || trbl.NeedDiePO_X6_X7ID == 9) // X6,7,x5
                        {
                            trbl.FinalStatusID = checkMRPOissue(trbl);

                        }
                        else
                        {
                            if (trbl.NeedDiePO_X6_X7ID == 6) // DMT repair
                            {
                                trbl.FinalStatusID = 6; //W-FA Result

                            }
                            else
                            {
                                trbl.FinalStatusID = 1; //Close
                            }
                        }
                        if (trbl.TroubleFrom.Contains("OUT"))
                        {
                            if (trbl.IsNeedPURConfirm == true)
                            {
                                trbl.isNeedRPA = false;
                            }
                            else
                            {
                                trbl.isNeedRPA = true;
                                sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);

                            }
                        }
                    }
                }
                goto Exit;
            }

            // AfterPE1App
            if (trbl.FinalStatusID == 14) // W-PE1-App
            {
                if (trbl.NeedDiePO_X6_X7ID == 1 || trbl.NeedDiePO_X6_X7ID == 2 || trbl.NeedDiePO_X6_X7ID == 7 || trbl.NeedDiePO_X6_X7ID == 8 || trbl.NeedDiePO_X6_X7ID == 9) // X7 or X6, x6OH, X9
                {
                    trbl.FinalStatusID = checkMRPOissue(trbl);
                }
                if (trbl.NeedDiePO_X6_X7ID == 3 || trbl.NeedDiePO_X6_X7ID == 6 || trbl.NeedDiePO_X6_X7ID == 11) // Maker Respone || DMT repair || Use spare
                {
                    trbl.FinalStatusID = 6; //W-FA-RESULT
                }

                if (trbl.NeedDiePO_X6_X7ID == 4 || trbl.NeedDiePO_X6_X7ID == 5) // No Need PO
                {
                    trbl.FinalStatusID = 1; // Close
                    trbl.CloseDate = today;
                    trbl.CloseContent = "No Repair die";
                }
                if (trbl.TroubleFrom.Contains("OUT"))
                {
                    if (trbl.IsNeedPURConfirm == true)
                    {
                        trbl.isNeedRPA = false;
                    }
                    else
                    {
                        trbl.isNeedRPA = true;
                        sendEmailJob.sendMailTPIorDFMNeedUploadToBOXToRPA(trbl, null);
                    }
                }
                goto Exit;
            }

        Exit:
            {
                // Update die
                var die = db.Die1.Find(trbl.DieID);
                if (die.DieClassify.Contains("MP"))
                {
                    die.DieStatusID = 3; // MP_Main
                }
                else
                {
                    var po = db.PO_Dies.Where(x => x.MR.TypeID < 4 && x.MR.DieNo == die.DieNo && x.Active != false && x.POStatusID != 20).FirstOrDefault();
                    if (po != null)
                    {
                        if (po.PaymentDate != null) // Đã paid
                        {
                            die.DieStatusID = 2; // W-MP
                        }
                        else
                        {
                            die.DieStatusID = 1; // Under_Making
                        }
                    }
                }
                die.RemarkDieStatusUsing = "Repair/Modify follow TPI: " + trbl.TroubleNo;
                die.Short = trbl.ActualShort;
                die.RecordDate = today;
                db.Entry(die).State = EntityState.Modified;
                db.SaveChanges();

                //DieStatusUpdateRegular newUpdate = new DieStatusUpdateRegular();
                //newUpdate.DieID = trbl.DieID;
                //newUpdate.DieNo = trbl.DieNo;
                //newUpdate.RecodeDate = today.ToString("yyyy-MM-dd");
                //newUpdate.ActualShort = trbl.ActualShort.Value.ToString();
                //newUpdate.DetailUsingStatus = " Repair/Modify follow TPI: " + trbl.TroubleNo;
                //newUpdate.DieOperationStatus = trbl.DetailPhenomenon;
                //newUpdate.StopDate = today.ToString("yyyy-MM-dd");
                //newUpdate.DieStatus = "Using";
                //newUpdate.IssueDate = today;
                //newUpdate.IssueBy = Session["Name"].ToString();
                //db.DieStatusUpdateRegulars.Add(newUpdate);
                //db.SaveChanges();
            }
            return trbl;
        }


        public JsonResult RPAConfirmDownloaded(int troubleID)
        {
            var status = false;
            if (Session["Name"].ToString().Contains("RPA"))
            {
                var trb = db.Troubles.Find(troubleID);
                trb.isNeedRPA = false;
                trb.RPAFinishedTime = DateTime.Now;
                db.Entry(trb).State = EntityState.Modified;
                db.SaveChanges();
                status = true;
            }
            return Json(status, JsonRequestBehavior.AllowGet);
        }

        public int checkMRPOissue(Trouble trbl)
        {
            var statusID = 0;

            if (trbl.MR_FinAppDate == null)
            {
                statusID = 4; //W-MR-Issue
            }
            else
            {
                if (trbl.PODate_PricePUS == null)
                {
                    statusID = 5; //W-PO-Issue
                }
                else
                {
                    if (trbl.FAResult == null || trbl.FAResult?.ToUpper() != "OK" || trbl.FAResult?.ToUpper() != "RS" || trbl.FAResult?.ToUpper() != "-")
                    {
                        statusID = 6; //W-FA-Result
                    }
                    else
                    {
                        statusID = 1; //W-FA-Result
                    }
                }
            }

            return statusID;
        }

        public ActionResult ExportExcel(List<Trouble> troubles)
        {
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/TroubleReport/FormTroubleList.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["Sheet1"];
                int rowId = 4;
                troubles = troubles.OrderByDescending(x => x.SubmitDate).ToList(); ;
                foreach (var trb in troubles)
                {

                    var existMR = db.MRs.Where(x => x.TroubleID.Contains(trb.TroubleID.ToString()) && x.StatusID != 11 && x.StatusID != 12 && x.Active != false).FirstOrDefault();
                    PO_Dies existPO = null;
                    if (existMR != null)
                    {
                        existPO = db.PO_Dies.Where(x => x.MRID == existMR.MRID).FirstOrDefault();
                    }

                    sheet.Cells["A" + rowId.ToString()].Value = trb.TroubleNo;
                    sheet.Cells["B" + rowId.ToString()].Value = trb.DieNo;
                    sheet.Cells["C" + rowId.ToString()].Value = trb.FinalStatusCalogory.FinalStatus;
                    sheet.Cells["D" + rowId.ToString()].Value = trb.Phase;
                    sheet.Cells["E" + rowId.ToString()].Value = trb.Progress;
                    sheet.Cells["F" + rowId.ToString()].Value = trb.TroubleFrom == null ? "" : trb.TroubleFrom;
                    sheet.Cells["G" + rowId.ToString()].Value = trb.TroubleLevelCalogory == null ? "" : trb.TroubleLevelCalogory.LevelType;
                    sheet.Cells["H" + rowId.ToString()].Value = trb.DeadLineIfUrgent;
                    sheet.Cells["I" + rowId.ToString()].Value = trb.DieTroubleRunningCalogory == null ? "" : trb.DieTroubleRunningCalogory.DieTroubelRunningStatus;
                    sheet.Cells["J" + rowId.ToString()].Value = trb.TroubleName;
                    sheet.Cells["K" + rowId.ToString()].Value = trb.Die1.ModelList.ModelName;
                    sheet.Cells["L" + rowId.ToString()].Value = db.Parts1.Where(x => x.PartNo == trb.Die1.PartNoOriginal).FirstOrDefault().Model;
                    sheet.Cells["M" + rowId.ToString()].Value = db.Suppliers.Where(x => x.SupplierName == trb.SupplierName).FirstOrDefault()?.SupplierCode;
                    sheet.Cells["N" + rowId.ToString()].Value = trb.Die1.ProcessCodeCalogory.Type;
                    sheet.Cells["O" + rowId.ToString()].Value = trb.HappenDate;
                    sheet.Cells["P" + rowId.ToString()].Value = trb.ActualStatusDetail;
                    sheet.Cells["Q" + rowId.ToString()].Value = trb.CheckingPoint;
                    sheet.Cells["R" + rowId.ToString()].Value = trb.ResponsibilityCalogory == null ? "" : trb.ResponsibilityCalogory.Responsibility;
                    sheet.Cells["S" + rowId.ToString()].Value = trb.NeedDiePOCalogory == null ? "" : trb.NeedDiePOCalogory.PODieCode;
                    sheet.Cells["T" + rowId.ToString()].Value = trb.TemapratyAction;
                    sheet.Cells["U" + rowId.ToString()].Value = trb.DetailPhenomenon;
                    sheet.Cells["V" + rowId.ToString()].Value = trb.DetailAction;

                    sheet.Cells["W" + rowId.ToString()].Value = commonFunction.getTroubleRootCause(trb);
                    sheet.Cells["X" + rowId.ToString()].Value = commonFunction.getTempoCM(trb);
                    sheet.Cells["Y" + rowId.ToString()].Value = commonFunction.getTempoCM(trb);
                    sheet.Cells["Z" + rowId.ToString()].Value = trb.TroubleReoccurCalogory == null ? "" : trb.TroubleReoccurCalogory.Reoccur;
                    sheet.Cells["AA" + rowId.ToString()].Value = trb.FeedbackForDieSpec;
                    sheet.Cells["AB" + rowId.ToString()].Value = trb.FeedbackForR_D;
                    sheet.Cells["AC" + rowId.ToString()].Value = trb.ActualShort;
                    sheet.Cells["AD" + rowId.ToString()].Value = trb.RenewDecisionCalogory == null ? "" : trb.RenewDecisionCalogory.DecisionStatus;

                    sheet.Cells["AE" + rowId.ToString()].Value = commonFunction.getTroubleType(trb);
                    sheet.Cells["AF" + rowId.ToString()].Value = trb.SubmitBy;
                    sheet.Cells["AG" + rowId.ToString()].Value = trb.SubmitDate;
                    sheet.Cells["AH" + rowId.ToString()].Value = trb.DMTCheckerComment;
                    sheet.Cells["AI" + rowId.ToString()].Value = trb.DMTCheckBy;
                    sheet.Cells["AJ" + rowId.ToString()].Value = trb.DMTCheckDate;
                    sheet.Cells["AK" + rowId.ToString()].Value = trb.DMTComment;
                    sheet.Cells["AL" + rowId.ToString()].Value = trb.DMTAppBy;
                    sheet.Cells["AM" + rowId.ToString()].Value = trb.DMTAppDate;
                    sheet.Cells["AN" + rowId.ToString()].Value = trb.PAEComment;
                    sheet.Cells["AO" + rowId.ToString()].Value = trb.PIC;
                    sheet.Cells["AP" + rowId.ToString()].Value = trb.FeedbackDate;
                    sheet.Cells["AQ" + rowId.ToString()].Value = trb.PAECommentBy;
                    sheet.Cells["AR" + rowId.ToString()].Value = trb.PAECommentDate;
                    sheet.Cells["AS" + rowId.ToString()].Value = trb.PE1Comment;
                    sheet.Cells["AT" + rowId.ToString()].Value = trb.NeedFA == true ? "Y" : "N";
                    sheet.Cells["AU" + rowId.ToString()].Value = trb.NeedTVP == true ? "Y" : "N";
                    sheet.Cells["AV" + rowId.ToString()].Value = trb.PE1CommentBy;
                    sheet.Cells["AW" + rowId.ToString()].Value = trb.PE1CommentDate;
                    sheet.Cells["AX" + rowId.ToString()].Value = existMR != null ? existMR.RequestDate.Value.ToShortDateString() : "";
                    sheet.Cells["AY" + rowId.ToString()].Value = existMR != null ? (existMR.PAEAppDate.HasValue ? existMR.PAEAppDate.Value.ToShortDateString() : "") : "";
                    sheet.Cells["AZ" + rowId.ToString()].Value = existPO != null ? (existPO.IssueDate.HasValue ? existPO.IssueDate.Value.ToShortDateString() : "Not Issue PO") : "";
                    sheet.Cells["BA" + rowId.ToString()].Value = existPO != null ? (existPO.PODate.HasValue ? existPO.PODate.Value.ToShortDateString() : "") : "";
                    sheet.Cells["BB" + rowId.ToString()].Value = trb.FASubmitDate;
                    sheet.Cells["BC" + rowId.ToString()].Value = trb.FAResult;
                    sheet.Cells["BD" + rowId.ToString()].Value = trb.CloseDate;
                    sheet.Cells["BE" + rowId.ToString()].Value = trb.CloseContent;
                    rowId++;
                }

                package.SaveAs(output);
            }
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=FORM_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx");
            output.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();
            return RedirectToAction("Index");
        }
        // Delete TPI
        public JsonResult Delete(int? id)
        {
            var status = false;
            if (Session["Admin"].ToString() == "Admin")
            {
                try
                {
                    var trb = db.Troubles.Find(id);
                    trb.Active = false;
                    db.Entry(trb).State = EntityState.Modified;
                    db.SaveChanges();
                    status = true;
                }
                catch
                {
                    status = false;
                }
            }

            return Json(status, JsonRequestBehavior.AllowGet);
        }
        public void convertToPDF(string fileName, int ID)
        {
            string path = Server.MapPath("~/File/TroubleReport/");
            string[] mainName = fileName.Split('.');
            string newName = mainName[0] + ".pdf";
            try
            {
                // Excel to PDF
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(path + fileName);
                // Save the document in PDF format

                workbook.Save(path + newName, SaveFormat.Pdf);
                var newFile = PDFEdit(newName);
                System.IO.File.Delete(path + newName);
                var trb = db.Troubles.Find(ID);
                trb.IsHandlePDF = true;
                db.Entry(trb).State = EntityState.Modified;
                db.SaveChanges();
            }
            catch
            {
                //Ppxt -> PDF
                //using (Presentation presentation = new Presentation(path + fileName))
                //{
                //    presentation.Save(path + newName, Aspose.Slides.Export.SaveFormat.Pdf);
                //}
            }

        }
        public string PDFEdit(string fileName)
        {
            //Document doc = new Document(PageSize.A4, 10, 10, 30, 1);
            //PdfWriter writer = PdfWriter.GetInstance(doc,new FileStream(Server.MapPath("~/Upload/" + awi.CtrlNo + "/" + awi.CtrlNo + ".pdf"), FileMode.Create));
            //doc.Open();
            //Image image = Image.GetInstance(Server.MapPath("~/Img/R2.png"));
            //image.SetAbsolutePosition(12, 300);
            //writer.DirectContent.AddImage(image, false);
            //doc.Close();

            string oldFile = Server.MapPath("~/File/TroubleReport/" + fileName);
            string newFile = Server.MapPath("~/File/TroubleReport/" + "[PDF]" + fileName);
            try
            {
                // open the reader
                PdfReader reader = new PdfReader(oldFile);
                Document document = new Document();

                // open the writer
                FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                // create the new page and add it to the pdf
                for (var i = 1; i <= reader.NumberOfPages; i++)
                {
                    var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    var importedPage = writer.GetImportedPage(reader, i);
                    document.SetPageSize(importedPage.BoundingBox);
                    document.NewPage();
                    var contentByte = writer.DirectContent;

                    contentByte.BeginText();
                    contentByte.SetFontAndSize(baseFont, 8);

                    var multiLineString = "PAGE: (" + Convert.ToString(i) + "/" + Convert.ToString(reader.NumberOfPages) + ")";


                    //Image image = Image.GetInstance(Server.MapPath("~/Img/R2.png"));
                    //image.ScalePercent(20);
                    //image.SetAbsolutePosition(/*importedPage.Width -*/ 120, importedPage.Height - 37);
                    contentByte.AddTemplate(importedPage, 0, 12);
                    //contentByte.AddImage(image, false);
                    contentByte.ShowTextAligned(PdfContentByte.ALIGN_LEFT, multiLineString, importedPage.Width - 60, 15, 0);
                    contentByte.EndText();

                }
                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();
            }
            catch
            {
                newFile = "";
            }


            return newFile;
        }


    }
}
