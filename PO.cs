using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using DMS03.Models;
using PagedList;
using System.IO;
using OfficeOpenXml;
using System.Net.Mail;
using System.Threading;
using System.Globalization;
using System.Web.UI.WebControls.WebParts;

namespace DMS03.Controllers
{
    public class PO_DiesController : Controller
    {
        private DMSEntities db = new DMSEntities();
        CommonFunctionController commonFunction = new CommonFunctionController();
        SendEmailController sendMailJob = new SendEmailController();

        // GET: PO_Dies
        public ActionResult Index(string search, int? page, string showAll, int? StatusID, string fromDate, string toDate, int? mRTypeID, string waitingFor, string export)
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            if (page == null) page = 1;
            int pageSize = 50;
            int pageNumber = (page ?? 1);
            var dept = Session["Dept"].ToString();
            var role = Session["PO_Role"].ToString();
            var grade = Session["Grade"] != null ? Session["Grade"].ToString() : "";
            var allPO = db.PO_Dies.Where(x => x.Active != false).ToList();
            List<PO_Dies> listPO = new List<PO_Dies>();
            // Nếu là PUR => show only PUR
            if (dept == "PUR" && role == "Check")
            {
                listPO = allPO.Where(x => x.POStatusID == 1 || x.POStatusID == 5 || x.POStatusID == 8 || x.POStatusID == 9 || x.POStatusID == 12 || x.POStatusID == 14 || x.POStatusID == 15 || x.POStatusID == 19 || x.POStatusID == 21 || x.POStatusID == 22).ToList();
            }
            if (dept == "PUR" && role == "Approve")
            {
                if (grade == "GM")
                {
                    listPO = allPO.Where(x => x.POStatusID == 10).ToList();
                }
                else
                {
                    listPO = allPO.Where(x => x.POStatusID == 2 || x.POStatusID == 11 || x.POStatusID == 16).ToList();
                }

            }
            // Nếu là PUC => Show onlu PUC
            if (dept == "PUC")
            {
                listPO = allPO.Where(x => x.POstatusCalogory.Status.Contains("PUC")).ToList();
            }

            // Nếu Là PUS => Show only PUS
            if (dept == "PUS" && role == "Check")
            {
                listPO = allPO.Where(x => x.POStatusID == 6 || x.POStatusID == 24).OrderByDescending(y => y.POStatusID).ToList();
            }
            if (dept == "PUS" && role == "Approve")
            {
                listPO = allPO.Where(x => x.POStatusID == 7).ToList();
            }

            // For Search
            if (!String.IsNullOrEmpty(search) || !String.IsNullOrEmpty(fromDate) || !String.IsNullOrEmpty(toDate) || mRTypeID > 0 || StatusID > 0)
            {
                var resultsearch = allPO;
                // Search string
                if (!String.IsNullOrEmpty(search))
                {
                    search = search.Trim();
                    // search số PO
                    //search PartNo
                    resultsearch = allPO.Where(x => x.MR.PartNo != null ? x.MR.PartNo.Contains(search) : x.POID == 0).ToList();

                    if (resultsearch.Count() == 0)
                    {
                        //search Status
                        resultsearch = allPO.Where(x => x.POStatusID == null ? x.POstatusCalogory.Status.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search PO Type
                        resultsearch = allPO.Where(x => x.MR.MRType != null ? x.MR.MRType.Type.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search Supplier Name
                        resultsearch = allPO.Where(x => x.MR.SupplierID != null ? x.MR.Supplier.SupplierName.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search Supplier Code
                        resultsearch = allPO.Where(x => x.MR.SupplierID != null ? x.MR.Supplier.SupplierCode.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search Die Make Code
                        resultsearch = allPO.Where(x => x.MR.OrderTo != null ? x.MR.OrderTo.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search MR No
                        resultsearch = allPO.Where(x => x.MR.MRNo != null ? x.MR.MRNo.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search Die No
                        resultsearch = allPO.Where(x => x.MR.DieNo != null ? x.MR.DieNo.Contains(search) : x.POID == 0).ToList();
                    }
                    if (resultsearch.Count() == 0)
                    {
                        //search IssueNO

                        resultsearch = allPO.Where(x => x.POIssueNo != null ? x.POIssueNo.Contains(search) : x.POID == 0).ToList();
                    }
                }
                // Search FromDate
                if (!String.IsNullOrEmpty(fromDate))
                {
                    resultsearch = resultsearch.Where(x => x.IssueDate >= Convert.ToDateTime(fromDate)).ToList();
                }
                // Search  toDate
                if (!String.IsNullOrEmpty(toDate))
                {
                    resultsearch = resultsearch.Where(x => x.IssueDate <= Convert.ToDateTime(toDate)).ToList();
                }
                // Search MRTypeID
                if (mRTypeID > 0)
                {
                    resultsearch = resultsearch.Where(x => x.MR.TypeID == mRTypeID).ToList();
                }
                // Search PO Status ID
                if (StatusID > 0)
                {
                    resultsearch = resultsearch.Where(x => x.POStatusID == StatusID).ToList();
                }

                listPO = resultsearch;
            }

            if (!String.IsNullOrEmpty(showAll))
            {
                listPO = db.PO_Dies.ToList();
            }
            //For Summary
            {
                // List For Waiting
                if (!String.IsNullOrEmpty(waitingFor))
                {
                    if (waitingFor == "w_PUR_Check")
                    {
                        var w_PUR_Check = allPO.Where(x => x.POStatusID == 1 || x.POStatusID == 9 || x.POStatusID == 15 || x.POStatusID == 21 || x.POStatusID == 22 && x.Active != false).ToList();
                        listPO = w_PUR_Check;
                    }
                    if (waitingFor == "w_PUR_App_NPIS")
                    {
                        var w_PUR_App_NPIS = allPO.Where(x => x.POStatusID == 5 || x.POStatusID == 14 || x.POStatusID == 19 && x.Active != false).ToList();
                        listPO = w_PUR_App_NPIS;
                    }
                    if (waitingFor == "w_PUR_M1")
                    {
                        var w_PUR_M1 = allPO.Where(x => x.POStatusID == 2 || x.POStatusID == 11 || x.POStatusID == 16 && x.Active != false).ToList();
                        listPO = w_PUR_M1;
                    }
                    if (waitingFor == "w_PUR_GM")
                    {
                        var w_PUR_GM = allPO.Where(x => x.POStatusID == 10 && x.Active != false).ToList();
                        listPO = w_PUR_GM;
                    }
                    if (waitingFor == "w_PUC_Input")
                    {
                        var w_PUC_Input = allPO.Where(x => x.POStatusID == 3 || x.POStatusID == 12 || x.POStatusID == 17 && x.Active != false).ToList();
                        listPO = w_PUC_Input;
                    }
                    if (waitingFor == "w_PUC_Dbc")
                    {
                        var w_PUC_Dbc = allPO.Where(x => x.POStatusID == 4 || x.POStatusID == 13 || x.POStatusID == 18 && x.Active != false).ToList();
                        listPO = w_PUC_Dbc;
                    }
                    if (waitingFor == "w_PUS_Check")
                    {
                        var w_PUS_Check = allPO.Where(x => (x.POStatusID == 6 || x.POStatusID == 24) && x.Active != false).ToList();
                        listPO = w_PUS_Check;
                    }
                    if (waitingFor == "w_PUS_App")
                    {
                        var w_PUS_App = allPO.Where(x => x.POStatusID == 7 && x.Active != false).ToList();
                        listPO = w_PUS_App;
                    }
                }

            }
            if (!String.IsNullOrEmpty(export))
            {
                exportPOToControlList(listPO);
            }
            ViewBag.w_PUR_Check_C = allPO.Where(x => x.POStatusID == 1 || x.POStatusID == 9 || x.POStatusID == 15 || x.POStatusID == 21 || x.POStatusID == 22 && x.Active != false).Count();
            ViewBag.w_PUR_App_NPIS_C = allPO.Where(x => x.POStatusID == 5 || x.POStatusID == 14 || x.POStatusID == 19 && x.Active != false).Count();
            ViewBag.w_PUR_M1_C = allPO.Where(x => x.POStatusID == 2 || x.POStatusID == 11 || x.POStatusID == 16 && x.Active != false).Count();
            ViewBag.w_PUR_GM_C = allPO.Where(x => x.POStatusID == 10 && x.Active != false).Count();
            ViewBag.w_PUC_Input_C = allPO.Where(x => x.POStatusID == 3 || x.POStatusID == 12 || x.POStatusID == 17 && x.Active != false).Count();
            ViewBag.w_PUC_Dbc_C = allPO.Where(x => x.POStatusID == 4 || x.POStatusID == 13 || x.POStatusID == 18 && x.Active != false).Count();
            ViewBag.w_PUS_Check_C = allPO.Where(x => x.POStatusID == 6 || x.POStatusID == 24 && x.Active != false).Count();
            ViewBag.w_PUS_App_C = allPO.Where(x => x.POStatusID == 7 && x.Active != false).Count();
            ViewBag.StatusID = new SelectList(db.POstatusCalogories, "POStatusID", "Status", StatusID);
            ViewBag.mRTypeID = new SelectList(db.MRTypes, "MR_ClassifyID", "Type", StatusID);
            ViewBag.search = search;
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;
            return View(listPO.Where(x => x.Active != false).OrderByDescending(x => x.CreateDate).ToPagedList(pageNumber, pageSize));
        }

        public JsonResult getPOdetail(int id)
        {
            //db.Configuration.ProxyCreationEnabled = false; // tránh lỗi vòng lặp

            var data = db.PO_Dies.Where(x => x.POID == id).AsEnumerable().Select(Po => new
            {
                PoIssueNo = Po.POIssueNo,
                PoStatus = Po.POstatusCalogory.Status,
                TemP = Po.TempPO == true ? "Y" : "N",
                MRNo = Po.MR.MRNo,
                VenderCode = Po.MR.Supplier.SupplierCode,
                VenderName = Po.MR.Supplier.SupplierName,
                DieMaker = Po.MR.OrderTo,
                DieNO = Po.MR.DieNo,
                PartNo = Po.MR.PartNo,
                PartName = Po.MR.PartName,
                Dim = Po.MR.Clasification,
                PR = Po.PR,
                Drawhis = Po.MR.DrawHis,
                ECNNo = Po.MR.ECNNo,
                CavQty = Po.MR.CavQty,
                PDD = Po.DeliveryDate.Value.ToString("yyyy-MM-dd"),
                DeleveryLocation = Po.DeliveryLocation,
                WarrantyShot = @String.Format("{0:0,0.##}", Po.WarrantyShot),
                TradeCoditionPName = Po.TradeConditionPName,
                ProductAbbrName = Po.MR.ModelName,
                UseBlockCode = Po.UseBlockCode,
                Currency = Po.MR.Unit,
                VendorFctry = Po.VendorFctry,
                NeedRegisterRateTableByPart = Po.NeedRegisterRateTableByPart,
                OrderQty = Po.OrderQty,
                ItemCategory = Po.ItemCategory,
                TransPortMethod = Po.TransportMethod,
                ContainerLoadingCode = Po.ContainerLoadingCode,
                Buyer = Po.BuyerCode,
                Remark = Po.Remark,
                ReasonIssuePO = Po.ReasonIssuePO,
                Price = @String.Format("{0:0,0.##}", Po.Price),
                EstimateCost = @String.Format("{0:0,0.##}", Po.MR.EstimateCost),
                POdate = Po.PODate == null ? "" : Po.PODate.Value.ToString("yyyy-MM-dd"),
                IssueBy = Po.IssueBy,
                IssueDate = Po.IssueDate == null ? "" : Po.IssueDate.Value.ToString("yyyy-MM-dd"),
                PURAppBy = Po.PURAppBy,
                PURAppDate = Po.PURAppDate == null ? "" : Po.PURAppDate.Value.ToString("yyyy-MM-dd"),
                PUCCheckBy = Po.PUCCheckBy,
                PUCCheckDate = Po.PUCCheckDate == null ? "" : Po.PUCCheckDate.Value.ToString("yyyy-MM-dd"),
                PUCAppBy = Po.PUCDoubleCheckBy,
                PUCAppDate = Po.PUCDoubleCheckDate == null ? "" : Po.PUCDoubleCheckDate.Value.ToString("yyyy-MM-dd"),
                PURAppNPISBy = Po.PURAppNPISBy,
                PURAppNPISDate = Po.PURAppNPISDate == null ? "" : Po.PURAppNPISDate.Value.ToString("yyyy-MM-dd"),
                PUSCheckBy = Po.PUSCheckBy,
                PUSCheckDate = Po.PUSCheckDate == null ? "" : Po.PUSCheckDate.Value.ToString("yyyy-MM-dd"),
                PUSAppBy = Po.PUSAppBy,
                PUSAppDate = Po.PUSAppDate == null ? "" : Po.PUSAppDate.Value.ToString("yyyy-MM-dd"),
                GMPURAppBy = Po.GMPURAppBy,
                GMPURAppDate = Po.GMPURAppDate == null ? "" : Po.GMPURAppDate.Value.ToString("yyyy-MM-dd"),
                //*******Change****************
                Change_OldDeliveryDate = Po.DeliveryDate == null ? "" : Po.DeliveryDate.Value.ToString("yyyy-MM-dd"),
                Change_NewDeliveryDate = Po.Change_DeliveryDate == null ? "" : Po.Change_DeliveryDate.Value.ToString("yyyy-MM-dd"),
                Change_Reason = Po.Change_Reason,
                Change_DeliveryKey = Po.Change_DeliveryKey,
                Change_RequestBy = Po.Change_RequestBy,
                Change_RequestDate = Po.Change_RequestDate == null ? "" : Po.Change_RequestDate.Value.ToString("yyyy-MM-dd"),
                Change_PURAppby = Po.Change_PURAppBy,
                Change_PURAppDate = Po.Change_PURAppDate == null ? "" : Po.Change_PURAppDate.Value.ToString("yyyy-MM-dd"),
                Change_PUCCheckBy = Po.Change_PUCCheckBy,
                Change_PUCCheckDate = Po.Change_PUCCheckDate == null ? "" : Po.Change_PUCCheckDate.Value.ToString("yyyy-MM-dd"),
                Change_PUCDoubleCheckBy = Po.Change_PUCDoubleCheckBy,
                Change_PUCDoubleCheckDate = Po.Change_PUCDoubleCheckDate == null ? "" : Po.Change_PUCDoubleCheckDate.Value.ToString("yyyy-MM-dd"),
                Change_PURAppNPISBy = Po.Change_PURAppNPISBy,
                Change_PURAppNPISDate = Po.Change_PURAppNPISDate == null ? "" : Po.Change_PURAppNPISDate.Value.ToString("yyyy-MM-dd"),
                change_Evidential = Po.Change_Evidential,
                //*******Cancel****
                cancel_Reason = Po.Cancel_Reason,
                Cancel_DeliveryKey = Po.Cancel_DeliveryKey,
                Cancel_RequestBy = Po.Cancel_RequestBy,
                Cancel_RequestDate = Po.Cancel_RequestDate == null ? "" : Po.Cancel_RequestDate.Value.ToString("yyyy-MM-dd"),
                Cancel_PURAppby = Po.Cancel_PURAppBy,
                Cancel_PURAppDate = Po.Cancel_PURAppDate == null ? "" : Po.Cancel_PURAppDate.Value.ToString("yyyy-MM-dd"),
                Cancel_PUCCheckBy = Po.Cancel_PUCCheckBy,
                Cancel_PUCCheckDate = Po.Cancel_PUCCheckDate == null ? "" : Po.Cancel_PUCCheckDate.Value.ToString("yyyy-MM-dd"),
                Cancel_PUCDoubleCheckBy = Po.Cancel_PUCDoubleCheckBy,
                Cancel_PUCDoubleCheckDate = Po.Cancel_PUCDoubleCheckDate == null ? "" : Po.Cancel_PUCDoubleCheckDate.Value.ToString("yyyy-MM-dd"),
                Cancel_PURAppNPISBy = Po.Cancel_PURAppNPISBy,
                Cancel_PURAppNPISDate = Po.Cancel_PURAppNPISDate == null ? "" : Po.Cancel_PURAppNPISDate.Value.ToString("yyyy-MM-dd"),
                cancel_Evidential = Po.Cancel_Evidential,

                ProcedureNo = Po.ProcedureNo,
                AttachNo_New = Po.AttachNo_New,
                AttachNo_Change = Po.AttachNo_Change,
                AttachNo_Cancel = Po.AttachNo_Cancel,
            });

            return Json(data, JsonRequestBehavior.AllowGet);
        }


        //***************** Issue PO New Area
        public JsonResult getPOforIssue(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();


            List<PO_Dies> listPO = new List<PO_Dies>();
            var poListAllowIssue = db.PO_Dies.Where(x => x.POStatusID == 1 || x.POStatusID == 9 || x.POStatusID == 21).ToList();
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR")
                {
                    var po = poListAllowIssue.Where(x => x.POID == Convert.ToInt32(id[i]));
                    listPO.AddRange(po);
                }
            }

            var data = listPO.Select(po => new
            {
                POID = po.POID,
                PartNo = po.MR.PartNo,
                Dim = po.MR.Clasification,
                VenderCode = po.MR.SupplierName,
                DieMakerCode = po.MR.OrderTo,
                PDD = po.DeliveryDate.Value.ToString("yyyy-MM-dd"),
            });


            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult issuePO(string[] id, string venderFctry, string NeedRegisterRateTableByPart, int OrderQty, string DeliveryLocation, int WarrantyShot, string TradeConditionPName, string UseBlockCode, string ItemCategory, string TransportMethod, string ContainerLoadingCode, string reasonIssuePO, string DeliveryDate)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            int success = 0;
            int fail = 0;
            List<PO_Dies> listPO = new List<PO_Dies>();
            var today = DateTime.Now;
            var poListAllowIssue = db.PO_Dies.Where(x => x.POStatusID == 1 || x.POStatusID == 9 || x.POStatusID == 21).ToList(); // Chỉ cho phép 3 loại nếu thêm thì sửa code tạo số PO
            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR")
                {
                    var po = poListAllowIssue.Find(x => x.POID == Convert.ToInt32(id[i]));
                    try
                    {
                        po.VendorFctry = venderFctry;
                        po.NeedRegisterRateTableByPart = NeedRegisterRateTableByPart;
                        po.OrderQty = OrderQty;
                        po.DeliveryLocation = DeliveryLocation;
                        po.WarrantyShot = WarrantyShot;
                        po.TradeConditionPName = TradeConditionPName;
                        po.UseBlockCode = UseBlockCode;
                        po.ItemCategory = ItemCategory;
                        po.TransportMethod = TransportMethod;
                        po.ContainerLoadingCode = ContainerLoadingCode;
                        po.ReasonIssuePO = reasonIssuePO;
                        po.IssueBy = Session["Name"].ToString();
                        po.IssueDate = today;
                        po.BuyerCode = Session["BuyerCode"].ToString();
                        po.ProcedureNo = currentProcedure.ProcedureNo;
                        po.AttachNo_New = currentProcedure.Att_NewPO;
                        DateTime pdd;
                        bool result = DateTime.TryParse(DeliveryDate, out pdd);
                        if (result)
                        {
                            po.DeliveryDate = pdd;
                            po.OriginalDeliveryDate = po.OriginalDeliveryDate != null ? po.OriginalDeliveryDate : pdd;
                        }
                        else
                        {
                            po.OriginalDeliveryDate = po.OriginalDeliveryDate != null ? po.OriginalDeliveryDate : po.DeliveryDate;
                        }
                        if (po.POStatusID == 1 || po.POStatusID == 9)
                        {
                            po.POIssueNo = createPOIssueNo(po.POIssueNo, true, false, false, false);
                        }
                        else // POStatusID = 21 Chú ý PO status thay đổi
                        {
                            po.POIssueNo = createPOIssueNo(po.POIssueNo, false, false, false, true);
                        }
                        po.POStatusID = 2;// W-PUR- approve

                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        commonFunction.genarateNewDie("", "", "", "", "", "", "", "", "", po.MR, po.WarrantyShot);
                        success = success + 1;

                    }
                    catch
                    {
                        fail = fail + 1;
                    }

                }
            }
            var status = new
            {
                success = success,
                fail = fail
            };
            return Json(status, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PurAPPNew(string[] id)
        {

            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            var grade = Session["Grade"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve" && grade == "M1")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 2 /*&& po.TempPO == false*/)
                    {
                        if (po.TempPO == true) // Chuyển cho GM PUR App
                        {
                            po.POStatusID = 10;// W-GM approve

                            sendMailJob.sendEmailTempPO("GMPUR", Session["Mail"].ToString(), po.MR.PartNo, po.MR.Clasification, po.MR.OrderTo, "Need GM Purcharsing Approval");
                        }
                        else // Chuyển cho PUC
                        {
                            po.POStatusID = 3; // W-PUC-Input-NPIS

                        }
                        po.PURAppBy = Session["Name"].ToString();
                        po.PURAppDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
                //For GM PUR trong truong hop temporary
                if (dept == "PUR" && role == "Approve" && grade == "GM")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 2 || po.POStatusID == 10)
                    {
                        po.POStatusID = 3; // W-PUC-Input-NPIS
                        po.GMPURAppBy = Session["Name"].ToString();
                        po.GMPURAppDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }

            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult purRejectNew(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            var grade = Session["Grade"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve" && grade == "M1")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 2)
                    {
                        po.POStatusID = 9; // Back-Wait-PUR-Issue
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
                //For GM PUR trong truong hop temporary
                if (dept == "PUR" && role == "Approve" && grade == "GM")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 2 || po.POStatusID == 10)
                    {
                        if (po.POStatusID == 2)
                        {
                            po.POStatusID = 9; // Back-Wait-PUR-Issue
                        }
                        else
                        {
                            po.POStatusID = 2; // Back cho M1 (W-PUR-M1-App)
                            sendMailJob.sendEmailTempPO("PURM1", Session["Mail"].ToString(), po.MR.PartNo, po.MR.Clasification, po.MR.OrderTo, " WAS REJECTED ");
                        }
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;

                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PUCDownload(string[] id)
        {

            List<PO_Dies> ListPO = new List<PO_Dies>();
            for (var i = 0; i < id.Length; i++)
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                ListPO.Add(po);
            }

            string handle = Guid.NewGuid().ToString();

            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/Format/FormPUCexportNewPO.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["FormPUCexportNewPO"];
                int rowId = 2;
                foreach (var po in ListPO)
                {
                    sheet.Cells["A" + rowId.ToString()].Value = 1;
                    sheet.Cells["B" + rowId.ToString()].Value = "VQ";
                    sheet.Cells["C" + rowId.ToString()].Value = po.MR.OrderTo;
                    sheet.Cells["D" + rowId.ToString()].Value = po.VendorFctry;
                    sheet.Cells["E" + rowId.ToString()].Value = po.MR.PartNo;
                    sheet.Cells["F" + rowId.ToString()].Value = po.MR.Clasification;
                    sheet.Cells["G" + rowId.ToString()].Value = po.PR;
                    sheet.Cells["H" + rowId.ToString()].Value = "5XX";
                    sheet.Cells["I" + rowId.ToString()].Value = "";
                    sheet.Cells["J" + rowId.ToString()].Value = po.MR.DrawHis;
                    sheet.Cells["K" + rowId.ToString()].Value = po.MR.ECNNo;
                    sheet.Cells["L" + rowId.ToString()].Value = "";
                    sheet.Cells["M" + rowId.ToString()].Value = String.Format("{0:yyyy/MM/dd}", po.DeliveryDate);
                    sheet.Cells["N" + rowId.ToString()].Value = "";
                    sheet.Cells["O" + rowId.ToString()].Value = 1;
                    sheet.Cells["P" + rowId.ToString()].Value = "SET";
                    sheet.Cells["Q" + rowId.ToString()].Value = "";
                    sheet.Cells["R" + rowId.ToString()].Value = po.DeliveryLocation;
                    sheet.Cells["S" + rowId.ToString()].Value = "N";
                    sheet.Cells["T" + rowId.ToString()].Value = po.UseBlockCode;
                    sheet.Cells["U" + rowId.ToString()].Value = "";
                    sheet.Cells["V" + rowId.ToString()].Value = "Y";
                    sheet.Cells["W" + rowId.ToString()].Value = 6;
                    sheet.Cells["X" + rowId.ToString()].Value = 1;
                    sheet.Cells["Y" + rowId.ToString()].Value = po.ItemCategory;
                    sheet.Cells["Z" + rowId.ToString()].Value = "";
                    sheet.Cells["AA" + rowId.ToString()].Value = "D";
                    sheet.Cells["AB" + rowId.ToString()].Value = "";
                    sheet.Cells["AC" + rowId.ToString()].Value = "";
                    sheet.Cells["AD" + rowId.ToString()].Value = 0;
                    sheet.Cells["AE" + rowId.ToString()].Value = po.TransportMethod;
                    sheet.Cells["AF" + rowId.ToString()].Value = po.TradeConditionPName == null ? "" : po.TradeConditionPName.Substring(0, 5);
                    sheet.Cells["AG" + rowId.ToString()].Value = po.TradeConditionPName == null ? "" : po.TradeConditionPName.Substring(0, 5);
                    sheet.Cells["AH" + rowId.ToString()].Value = "CVN";
                    sheet.Cells["AI" + rowId.ToString()].Value = po.ContainerLoadingCode;
                    sheet.Cells["AJ" + rowId.ToString()].Value = "";
                    sheet.Cells["AK" + rowId.ToString()].Value = po.MR.CavQty;
                    sheet.Cells["AL" + rowId.ToString()].Value = po.WarrantyShot;
                    sheet.Cells["AM" + rowId.ToString()].Value = po.MR.Clasification;
                    sheet.Cells["AN" + rowId.ToString()].Value = "";
                    sheet.Cells["AO" + rowId.ToString()].Value = 1;
                    sheet.Cells["AP" + rowId.ToString()].Value = "";
                    sheet.Cells["AQ" + rowId.ToString()].Value = "";
                    sheet.Cells["AR" + rowId.ToString()].Value = 0;
                    sheet.Cells["AS" + rowId.ToString()].Value = "";
                    sheet.Cells["AT" + rowId.ToString()].Value = po.MR.ModelName.Length > 15 ? po.MR.ModelName.Substring(0, 15) : po.MR.ModelName;
                    sheet.Cells["AU" + rowId.ToString()].Value = 9100;
                    sheet.Cells["AV" + rowId.ToString()].Value = po.BuyerCode;
                    sheet.Cells["AW" + rowId.ToString()].Value = po.BuyerCode;
                    sheet.Cells["AX" + rowId.ToString()].Value = ""; //orgin ctry
                    sheet.Cells["AY" + rowId.ToString()].Value = po.MR.PartName;
                    sheet.Cells["AZ" + rowId.ToString()].Value = po.MR.MRNo;
                    sheet.Cells["BA" + rowId.ToString()].Value = "";
                    sheet.Cells["BB" + rowId.ToString()].Value = "045M01";
                    rowId++;
                }

                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }
            //Response.Clear();
            //Response.Buffer = true;
            //Response.Charset = "";
            //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("content-disposition", "attachment;filename=PO_DIe" + ".csv");
            //output.WriteTo(Response.OutputStream);
            //Response.Flush();
            //Response.End();
            var data = new { FileGuid = handle, FileName = "PO_DIe.csv" };
            return Json(data, JsonRequestBehavior.AllowGet);

        }
        public JsonResult pucConfirm(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Check" || role == "Approve"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 3)
                    {
                        po.POStatusID = 4; // W-PUS-double Check
                        po.PUCCheckBy = Session["Name"].ToString();
                        po.PUCCheckDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult doubleCheckNew(string id, HttpPostedFileBase file)
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            List<string> suscussPONO = new List<string>();
            if (dept == "PUC" && role == "Approve")
            {
                var listPO = db.PO_Dies.Where(x => x.POStatusID == 4).ToList();
                // Luu file Execl
                string fileName = "FileDoubleCheckNew-" + today.ToString("yyyy-MM-dd-hhss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/PO/");
                fileName += fileExt;
                file.SaveAs(path + Path.GetFileName(fileName));
                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    for (int row = start.Row + 1; row <= end.Row; row++)
                    { // Row by row...
                        var partNo = worksheet.Cells[row, 7].Text.Trim();
                        var drawHis = worksheet.Cells[row, 12].Text.Trim();
                        if (drawHis.Length == 1)
                        {
                            drawHis = "00" + drawHis;
                        }
                        if (drawHis.Length == 2)
                        {
                            drawHis = "0" + drawHis;
                        }
                        var ECN = worksheet.Cells[row, 13].Text.Trim();
                        if (partNo == null) break;
                        var dim = worksheet.Cells[row, 8].Text.Trim();
                        var ExistPO = listPO.Where(x => x.MR.PartNo.Trim() == partNo.Trim() && x.MR.Clasification.Trim() == dim.Trim() && x.MR.DrawHis == drawHis && x.MR.ECNNo == ECN && x.Active != false).FirstOrDefault();
                        if (ExistPO == null)
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi: Sai draw his hoặc sai số ECN hoặc status PO này chưa thể double check");
                            goto exitLoop;
                        }
                        //****
                        var vendercode = worksheet.Cells[row, 5].Text.Trim();
                        if (vendercode != ExistPO.MR.OrderTo.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi vender Code Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var venderFctry = worksheet.Cells[row, 6].Text.Trim();
                        if (venderFctry.Length == 1)
                        {
                            venderFctry = "0" + venderFctry;
                        }
                        if (venderFctry != ExistPO.VendorFctry.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi venderFctry Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var PR = worksheet.Cells[row, 9].Text.Trim();
                        if (PR != ExistPO.PR.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi PR Không chính xác");
                            goto exitLoop;
                        }

                        //****


                        ////****
                        var deliveryDate = Convert.ToDateTime(worksheet.Cells[row, 15].Text.Trim()).ToShortDateString();
                        if (deliveryDate != ExistPO.DeliveryDate.Value.ToShortDateString())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Delivery Date Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var orderQty = worksheet.Cells[row, 18].Text.Trim();
                        if (orderQty != ExistPO.OrderQty.ToString().Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi OderQty Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var itemCategory = worksheet.Cells[row, 29].Text.Trim();
                        if (itemCategory != ExistPO.ItemCategory.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi item Category Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var Transport = worksheet.Cells[row, 42].Text.Trim();
                        if (Transport != ExistPO.TransportMethod.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Transport Method Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var departurePort = worksheet.Cells[row, 43].Text.Trim();
                        if (departurePort != ExistPO.TradeConditionPName.Substring(0, 5))
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Departure Port Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var arrival = worksheet.Cells[row, 44].Text.Trim();
                        if (arrival != ExistPO.TradeConditionPName.Substring(0, 5))
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Arrival Port Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var WithDrawQty = worksheet.Cells[row, 50].Text.Trim();
                        if (WithDrawQty != ExistPO.MR.CavQty.ToString())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi With Draw Qty (Cav qty) Không chính xác");
                            goto exitLoop;
                        }
                        //****
                        var warranty = worksheet.Cells[row, 51].Text.Trim();
                        if (warranty != ExistPO.WarrantyShot.ToString())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi warranty Không chính xác");
                            goto exitLoop;
                        }

                        var DieNo = worksheet.Cells[row, 52].Text.Trim();
                        if (DieNo != ExistPO.MR.Clasification)
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Die Number Không chính xác");
                            goto exitLoop;
                        }


                        //// TAM THỜI BỎ CHECK
                        //var ProductAbbrName = worksheet.Cells[row, 55].Text.Trim();
                        //ProductAbbrName = ProductAbbrName.Length > 8 ? ProductAbbrName.Substring(0, 8) : ProductAbbrName;
                        //var modelName = ExistPO.MR.ModelName.Length > 8 ? ExistPO.MR.ModelName.Substring(0, 8) : ExistPO.MR.ModelName;
                        //if (ProductAbbrName != modelName)
                        //{
                        //    fail.Add(partNo + "-" + dim + "Lỗi Product Abbr Name Không chính xác");
                        //    goto exitLoop;
                        //}

                        //var QuotationPIC = worksheet.Cells[row, 59].Text.Trim();
                        //if (QuotationPIC != ExistPO.BuyerCode.Trim())
                        //{
                        //    fail.Add(partNo + "-" + dim + "Lỗi Quotation PIC Không chính xác");
                        //    goto exitLoop;
                        //}

                        //var buyer = worksheet.Cells[row, 60].Text.Trim();
                        //if (buyer != ExistPO.BuyerCode.Trim())
                        //{
                        //    fail.Add(partNo + "-" + dim + "Lỗi Buyer Không chính xác");
                        //    goto exitLoop;
                        //}

                        // Xử lí khi check xong ko có lỗi
                        // 1. Luu PO_New


                        // 2.Thay đổi trạng thái PO status
                        ExistPO.POStatusID = 5; // W-PUR-App-NPIS
                        ExistPO.PUCDoubleCheckBy = Session["Name"].ToString();
                        ExistPO.PUCDoubleCheckDate = today;
                        db.Entry(ExistPO).State = EntityState.Modified;
                        db.SaveChanges();
                        if (ExistPO.TempPO == true)
                        {
                            sendMailJob.sendEmailTempPO("PUS", Session["Mail"].ToString(), ExistPO.MR.PartNo, ExistPO.MR.Clasification, ExistPO.MR.OrderTo, "Need PUS input price in NPIS");
                        }
                        suscussPONO.Add(ExistPO.POIssueNo);
                        success.Add(partNo + "-" + dim + "OK");

                        // Luu vào PO_New
                        {
                            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
                            PO_New PoNew = new PO_New();
                            PoNew.POID = ExistPO.POID;
                            PoNew.MRID = ExistPO.MRID;
                            PoNew.POIssueNo = ExistPO.POIssueNo;
                            PoNew.TempPO = ExistPO.TempPO;
                            PoNew.PR = ExistPO.PR;
                            PoNew.DeliveryDate = ExistPO.DeliveryDate;
                            PoNew.DeliveryLocation = ExistPO.DeliveryLocation;
                            PoNew.WarrantyShot = ExistPO.WarrantyShot;
                            PoNew.TradeConditionPName = ExistPO.TradeConditionPName;
                            PoNew.UseBlockCode = ExistPO.UseBlockCode;
                            PoNew.VendorFctry = ExistPO.VendorFctry;
                            PoNew.NeedRegisterRateTableByPart = ExistPO.NeedRegisterRateTableByPart;
                            PoNew.OrderQty = ExistPO.OrderQty;
                            PoNew.ItemCategory = ExistPO.ItemCategory;
                            PoNew.TransportMethod = ExistPO.TransportMethod;
                            PoNew.ContainerLoadingCode = ExistPO.ContainerLoadingCode;
                            PoNew.BuyerCode = ExistPO.BuyerCode;
                            PoNew.Price = ExistPO.Price;
                            PoNew.PODate = ExistPO.PODate;
                            PoNew.IssueBy = ExistPO.IssueBy;
                            PoNew.IssueDate = ExistPO.IssueDate;
                            PoNew.PURAppBy = ExistPO.PURAppBy;
                            PoNew.PURAppDate = ExistPO.PURAppDate;
                            PoNew.PUCCheckBy = ExistPO.PUCCheckBy;
                            PoNew.PUCCheckDate = ExistPO.PUCCheckDate;
                            PoNew.PUCDoubleCheckBy = ExistPO.PUCDoubleCheckBy;
                            PoNew.PUCDoubleCheckDate = ExistPO.PUCDoubleCheckDate;
                            PoNew.PURAppNPISBy = ExistPO.PURAppNPISBy;
                            PoNew.PURAppNPISDate = ExistPO.PURAppNPISDate;
                            PoNew.PUSCheckBy = ExistPO.PUSCheckBy;
                            PoNew.PUSCheckDate = ExistPO.PUSCheckDate;
                            PoNew.PUSAppBy = ExistPO.PUSAppBy;
                            PoNew.PUSAppDate = ExistPO.PUSAppDate;
                            PoNew.PaymentDate = ExistPO.PaymentDate;
                            PoNew.ProcedureNo = currentProcedure.ProcedureNo;
                            PoNew.AttachmentNo = currentProcedure.Att_NewPO;
                            PoNew.Active = true;
                            PoNew.GMPURAppBy = ExistPO.GMPURAppBy;
                            PoNew.GMPURAppDate = ExistPO.GMPURAppDate;
                            db.PO_New.Add(PoNew);
                            db.SaveChanges();

                        }



                    exitLoop:
                        ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    }
                }
            }
            var data = new
            {
                suscussPONO,
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pucReject(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For reject
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Approve" || role == "Check"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 3 || po.POStatusID == 4)
                    {
                        po.POStatusID = 9; // Back-Wait-PUR-Issue
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " Was Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PURApproveNPISNew(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Check")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 5)
                    {
                        if (po.PUSAppDate != null)
                        {
                            if (po.Price == 10)
                            {
                                po.POStatusID = 24; // W-PUS-Input price (Actual)
                            }
                            else
                            {
                                if (po.PaymentDate == null)
                                {
                                    po.POStatusID = 8; // W-Payment
                                }
                                else
                                {
                                    po.POStatusID = 23; // Paid
                                }
                            }
                           
                        }
                        else
                        {

                            po.POStatusID = 6; // W-PUS-Input price
                        }


                        po.PURAppNPISBy = Session["Name"].ToString();
                        po.PURAppNPISDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PUSinputPrice(string[] id, string price, string poDate)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for update price 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUS" && (role == "Check" || role == "Approve"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 6 || po.POStatusID == 7 || po.POStatusID == 8 || po.POStatusID == 24)
                    {
                        po.POStatusID = 7; // W-PUS-Approve
                        po.Price = Convert.ToDouble(price);
                        po.PODate = Convert.ToDateTime(poDate);
                        po.OriginalPOdate = po.OriginalPOdate != null ? po.OriginalPOdate : po.PODate;
                        po.PUSCheckBy = Session["Name"].ToString();
                        po.PUSCheckDate = DateTime.Now;


                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                        commonFunction.genarateNewDie("", "", "", "", "", "", "", "", "", po.MR, po.WarrantyShot);

                        if (po.TempPO == true && po.Price == 10)
                        {
                            sendMailJob.sendEmailTempPO("PUSAPP", Session["Mail"].ToString(), po.MR.PartNo, po.MR.Clasification, po.MR.OrderTo, "Need PUS Approval Temporary Price.");
                        }
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);

        }
        public JsonResult pusApprove(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for update price 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUS" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 7)
                    {
                        if (po.Price == 10 && po.MR.EstimateCost == 10)
                        {
                            po.POStatusID = 24; // W-PUS-INPUT-PRICE(Actual)
                        }
                        else
                        {
                            if (po.PaymentDate == null)
                            {
                                po.POStatusID = 8; // W-Payment
                            }
                            else
                            {
                                po.POStatusID = 23; // Paid
                            }

                        }

                        // XỬ LÍ MR KHI GIÁ ĐƯỢC APP
                        //1. xử lí exchange tiền tệ 
                        var mR = db.MRs.Find(po.MRID);
                        mR.AppCost = po.Price;
                        mR.PODate = po.PODate;
                        // Kiểm tra Unit to exchange tiền tệ
                        var resultExchange = commonFunction.exchangeToUSD(mR.AppCost, mR.Unit);
                        mR.ExchangeRate = resultExchange.rate;
                        mR.AppCostExchangeUSD = resultExchange.price;
                        //if (mR.Unit == "USD")
                        //{
                        //    //mR.ExchangeRate = null;
                        //    mR.AppCostExchangeUSD = mR.AppCost;
                        //}
                        //else
                        //{
                        //    if (mR.Unit == "VND")
                        //    {
                        //        mR.ExchangeRate = db.ExchangeRates.ToList().LastOrDefault().RateVNDtoUSD;
                        //        mR.AppCostExchangeUSD = System.Math.Round(Convert.ToDouble(mR.AppCost / mR.ExchangeRate), 2);
                        //    }
                        //    else
                        //    if (mR.Unit == "JPY")
                        //    {
                        //        mR.ExchangeRate = db.ExchangeRates.ToList().LastOrDefault().RateJPYtoUSD;
                        //        mR.AppCostExchangeUSD = System.Math.Round(Convert.ToDouble(mR.AppCost / mR.ExchangeRate), 2);
                        //    }
                        //}

                        // *********Xong xử lí đổi tiền tệ
                        //************************************************
                        //2. Xử lí TemPO => y/c PLAN update budget code
                        //if (po.TempPO == true && mR.AppCostExchangeUSD * db.ExchangeRates.ToList().Last().RateVNDtoUSD > 30000000 && mR.TypeID != 6)
                        //{
                        //    mR.EstimateCost = mR.AppCost;
                        //    mR.ReUpdateBudgetCode = true;
                        //    sendEmailToPlanToReupdateBudget(mR.MRNo);
                        //}

                        bool isNeedPlan_estCost = commonFunction.isNeedPLAN((int)mR.TypeID, (double)mR.EstimateCost, mR.Unit);
                        bool isNeedPlan_appCost = commonFunction.isNeedPLAN((int)mR.TypeID, (double)mR.AppCost, mR.Unit);
                        if (isNeedPlan_estCost != isNeedPlan_appCost)
                        {
                            mR.EstimateCost = mR.EstimateCost == 10 ? mR.AppCost : mR.EstimateCost;
                            mR.ReUpdateBudgetCode = true;
                            sendMailJob.sendEmailToPlanToReupdateBudget(mR.MRNo);
                        }

                        if (po.TempPO == true && po.Price > 10)
                        {
                            mR.EstimateCost = mR.AppCost;
                            mR.EstimateCostExchangeUSD = mR.AppCostExchangeUSD;
                        }
                        //Luu MR
                        mR.StatusID = 13;
                        // Check lại và auto fill Acc info
                        if (mR.Belong != "CRG")
                        {
                            var modelColorOrMono = db.ModelLists.Find(mR.ModelID).ModelType;
                            var supplierCode = db.Suppliers.Find(mR.SupplierID).SupplierCode;
                            var first2LeterPartNo = mR.PartNo.Remove(2, mR.PartNo.Length - 2);
                            string[] accInfor = commonFunction.AutoAccFill(mR.TypeID.Value, modelColorOrMono, supplierCode, mR.AppCost.Value, mR.Unit, first2LeterPartNo);
                            mR.GLAccount = accInfor[0];
                            mR.Location = accInfor[1];
                            mR.AssetNumber = accInfor[2];
                        }

                        db.Entry(mR).State = EntityState.Modified;


                        //3. Xu li khi la new die X1,X4
                        if (mR.TypeID == 1 || mR.TypeID == 2 || mR.TypeID == 3) // 1: new: 2: addition; 3: Renewal
                        {
                            var newDieID = commonFunction.genarateNewDie(null, null, null, null, null, null, null, null, dept, mR, po.WarrantyShot);
                        }
                        // 4. Xu li close trouble
                        commonFunction.UpdateTPIStatus(mR, "POIssue");


                        if (!String.IsNullOrEmpty(mR.SucessDieID))
                        {
                            if (mR.SucessDieID.Trim().Length == 20)
                            {
                                handelSuccesDie(mR.MRID);
                            }

                        }

                        if (!String.IsNullOrEmpty(mR.CommonPart) && String.IsNullOrEmpty(mR.SucessDieID))
                        {
                            if (mR.CommonPart.Trim().Length > 7)
                            {
                                handelCommonOrFaminlyPart(mR.DieNo, mR.CommonPart, mR.ModelName);
                            }
                        }
                        if (!String.IsNullOrEmpty(mR.FamilyPart) && String.IsNullOrEmpty(mR.SucessDieID))
                        {
                            if (mR.FamilyPart.Trim().Length > 7)
                            {
                                handelCommonOrFaminlyPart(mR.DieNo, mR.FamilyPart, mR.ModelName);
                            }
                        }

                        //Luu PO 
                        po.PUSAppBy = Session["Name"].ToString();
                        po.PUSAppDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        //updateDieLaunching(po);

                        success.Add(po.POIssueNo);


                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pusReject(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for update price 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUS" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 7)
                    {
                        po.POStatusID = 6; // W-PUS-Input price
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " was rejected: " + reason + System.Environment.NewLine + po.Remark;

                        //Luu PO 
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        //updateDieLaunching(po);
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult ExportForRequestPOForm(string[] id)
        {
            List<PO_Dies> ListPO = new List<PO_Dies>();
            for (var i = 0; i < id.Length; i++)
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                ListPO.Add(po);
            }
            string handle = Guid.NewGuid().ToString();
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/Format/PO_Form.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["PO"];
                sheet.Cells["T6"].Value = Session["Name"].ToString();
                sheet.Cells["AA6"].Value = DateTime.Now.ToString("yyyy-MMM-dd");
                sheet.Cells["AA7"].Value = DateTime.Now.ToString("yyyy-MMM-dd");

                int rowId = 15;
                int No = 1;
                foreach (var po in ListPO)
                {
                    sheet.Cells["B" + rowId.ToString()].Value = No;
                    sheet.Cells["C" + rowId.ToString()].Value = po.MR.OrderTo;
                    sheet.Cells["D" + rowId.ToString()].Value = po.VendorFctry;
                    sheet.Cells["E" + rowId.ToString()].Value = po.MR.Unit;
                    sheet.Cells["F" + rowId.ToString()].Value = po.NeedRegisterRateTableByPart;
                    sheet.Cells["G" + rowId.ToString()].Value = po.MR.PartNo;
                    sheet.Cells["H" + rowId.ToString()].Value = po.MR.Clasification;
                    sheet.Cells["I" + rowId.ToString()].Value = po.MR.PartName;
                    sheet.Cells["J" + rowId.ToString()].Value = po.PR;
                    sheet.Cells["K" + rowId.ToString()].Value = po.MR.DrawHis;
                    sheet.Cells["L" + rowId.ToString()].Value = po.MR.ECNNo;
                    sheet.Cells["M" + rowId.ToString()].Value = po.DeliveryDate == null ? "" : po.DeliveryDate.Value.ToString("yyyy-MMM-dd");
                    sheet.Cells["N" + rowId.ToString()].Value = po.OrderQty;
                    sheet.Cells["O" + rowId.ToString()].Value = po.DeliveryLocation;
                    sheet.Cells["P" + rowId.ToString()].Value = po.UseBlockCode;
                    sheet.Cells["Q" + rowId.ToString()].Value = po.ItemCategory;
                    sheet.Cells["R" + rowId.ToString()].Value = po.TransportMethod;
                    sheet.Cells["S" + rowId.ToString()].Value = po.ContainerLoadingCode;
                    sheet.Cells["T" + rowId.ToString()].Value = po.MR.CavQty;
                    sheet.Cells["U" + rowId.ToString()].Value = po.WarrantyShot;
                    sheet.Cells["V" + rowId.ToString()].Value = po.TradeConditionPName;
                    sheet.Cells["Y" + rowId.ToString()].Value = po.MR.MRNo;
                    sheet.Cells["Z" + rowId.ToString()].Value = po.MR.ModelName.Length > 8 ? po.MR.ModelName.Substring(0, 8) : po.MR.ModelName;
                    sheet.Cells["AA" + rowId.ToString()].Value = po.BuyerCode;
                    rowId++;
                    No++;
                }
                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }
            var data = new { FileGuid = handle, FileName = "PO_Request_File.xlsx" };
            return Json(data, JsonRequestBehavior.AllowGet);
        }
        // ************KẾT THÚC PO NEW



        // ************CHANGE PO Area
        // Nguyên tắc: 
        // 1.Chỉ cho phép change PO khi PO đó đang ở trạng thái W-Payment : POStatusID = 8
        // 2. Sẽ ghi nhập PO CHange sau khi PUR App NPIS => POStatus = 14 Change_W-PUR-APP NPIS
        public JsonResult checkAllowChang(string[] id)
        {

            List<object> result = new List<object>();
            for (var i = 0; i < id.Length; i++)
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
                if (po.POstatusCalogory.isCanChangePO == true)
                {
                    var data = new
                    {
                        POID = po.POID,
                        PartNo = po.MR.PartNo,
                        Dim = po.MR.Clasification,
                        VenderCode = po.MR.SupplierName,
                        DieMakerCode = po.MR.OrderTo,
                        PDD = po.DeliveryDate.Value.ToString("yyyy-MM-dd"),
                        ProcedureNo = currentProcedure.ProcedureNo,
                        AttachNo = currentProcedure.Att_ChangePO
                    };
                    result.Add(data);
                }
            }


            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public JsonResult ChangePO(string[] id, string new_deliveryDate, string change_DeliveryKey, string change_reason, HttpPostedFileBase file)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();

            List<string> success = new List<string>();
            List<PO_Dies> listPO = new List<PO_Dies>();
            var today = DateTime.Now;
            var po = db.PO_Dies.Find(int.Parse(id[0]));
            if (dept == "PUR" && po.POstatusCalogory.isCanChangePO == true) // Cho phep change PO khi W-PAyemt (8) || Bi Back lai (15)
            {
                po.POStatusID = 11; // Change_W-M1-PUR- App
                po.Change_DeliveryDate = Convert.ToDateTime(new_deliveryDate);

                //if (!String.IsNullOrEmpty(change_DeliveryKey))
                //{
                //    if (change_DeliveryKey.Length == 6)
                //    {
                //        change_DeliveryKey = change_DeliveryKey + "0000";
                //    }
                //    if (change_DeliveryKey.Length == 7)
                //    {
                //        change_DeliveryKey = change_DeliveryKey + "000";
                //    }
                //    if (change_DeliveryKey.Length == 8)
                //    {
                //        change_DeliveryKey = change_DeliveryKey + "00";
                //    }
                //    if (change_DeliveryKey.Length == 9)
                //    {
                //        change_DeliveryKey = change_DeliveryKey + "0";
                //    }
                //}

                po.Change_DeliveryKey = change_DeliveryKey;
                po.Change_Reason = change_reason;
                po.ProcedureNo = currentProcedure.ProcedureNo;
                po.AttachNo_Change = currentProcedure.Att_ChangePO;
                if (file != null) // Luu file
                {
                    string fileName = "Evidential-" + today.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/PO/Evidential/");
                    fileName += fileExt;
                    file.SaveAs(path + Path.GetFileName(fileName));
                    po.Change_Evidential = fileName;
                }
                // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                var oldPOIssueNo = po.POIssueNo;
                // Tạo số PO mới
                po.POIssueNo = createPOIssueNo(po.POIssueNo, false, true, false, false);
                po.Change_RequestBy = Session["Name"].ToString();
                po.Change_RequestDate = today;
                // Luu DB
                db.Entry(po).State = EntityState.Modified;
                db.SaveChanges();
                success.Add(oldPOIssueNo);
            }
            var status = new
            {
                success = success,

            };
            return Json(status, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PurAPPChange(string[] id)
        {

            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1 || GM
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 11) // Change_W-PUR-M1-App
                    {
                        po.POStatusID = 12; // Change_W-PUC-Input-NPIS
                        po.Change_PURAppBy = Session["Name"].ToString();
                        po.Change_PURAppDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }

            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pucConfirmChange(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Check" || role == "Approve"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 12) //Change_W-PUC-Input NPIS
                    {
                        po.POStatusID = 13; // Change_W-PUC-double Check
                        po.Change_PUCCheckBy = Session["Name"].ToString();
                        po.Change_PUCCheckDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);

                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PUCDownloadChange_Cancel(string[] id)
        {

            List<PO_Dies> ListPO = new List<PO_Dies>();
            for (var i = 0; i < id.Length; i++)
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                ListPO.Add(po);
            }

            string handle = Guid.NewGuid().ToString();

            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/Format/FormPUCexportChangeAndCancel.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["Change_Cancel"];
                int rowId = 2;
                foreach (var po in ListPO)
                {
                    var check = "";
                    var deliveryKey = "";
                    if (po.POStatusID == 12 || po.POStatusID == 13)
                    {
                        check = "R";
                        deliveryKey = po.Change_DeliveryKey;
                    }
                    if (po.POStatusID == 17 || po.POStatusID == 18)
                    {
                        check = "C";
                        deliveryKey = po.Cancel_DeliveryKey;
                    }
                    try
                    {
                        deliveryKey = deliveryKey.Substring(0, 9);
                    }
                    catch
                    {
                        //
                    }

                    sheet.Cells["A" + rowId.ToString()].Value = check;
                    sheet.Cells["B" + rowId.ToString()].Value = "VQ";
                    sheet.Cells["C" + rowId.ToString()].Value = po.MR.OrderTo;
                    sheet.Cells["D" + rowId.ToString()].Value = po.MR.PartNo;
                    sheet.Cells["E" + rowId.ToString()].Value = deliveryKey;
                    sheet.Cells["F" + rowId.ToString()].Value = "+001";
                    sheet.Cells["G" + rowId.ToString()].Value = "";
                    sheet.Cells["H" + rowId.ToString()].Value = "";
                    sheet.Cells["I" + rowId.ToString()].Value = check == "R" ? String.Format("{0:yyyy/MM/dd}", po.Change_DeliveryDate) : "";
                    sheet.Cells["J" + rowId.ToString()].Value = "";
                    sheet.Cells["K" + rowId.ToString()].Value = po.OrderQty;
                    sheet.Cells["L" + rowId.ToString()].Value = po.DeliveryLocation;
                    sheet.Cells["M" + rowId.ToString()].Value = po.UseBlockCode;
                    sheet.Cells["N" + rowId.ToString()].Value = "";
                    sheet.Cells["O" + rowId.ToString()].Value = po.TransportMethod;
                    sheet.Cells["P" + rowId.ToString()].Value = po.TradeConditionPName;
                    sheet.Cells["Q" + rowId.ToString()].Value = po.TradeConditionPName;
                    sheet.Cells["R" + rowId.ToString()].Value = po.ContainerLoadingCode;
                    sheet.Cells["S" + rowId.ToString()].Value = "";
                    sheet.Cells["T" + rowId.ToString()].Value = "";
                    rowId++;
                }
                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }
            //Response.Clear();
            //Response.Buffer = true;
            //Response.Charset = "";
            //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //Response.AddHeader("content-disposition", "attachment;filename=PO_DIe" + ".csv");
            //output.WriteTo(Response.OutputStream);
            //Response.Flush();
            //Response.End();
            var data = new { FileGuid = handle, FileName = "Change+Cancel.csv" };
            return Json(data, JsonRequestBehavior.AllowGet);

        }
        public JsonResult doubleCheckChange(string id, HttpPostedFileBase file)
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            List<string> suscussPONO = new List<string>();
            if (dept == "PUC" && role == "Approve")
            {

                var listPO = db.PO_Dies.Where(x => x.POStatusID == 13 && x.Active != false).ToList();

                // Luu file Execl
                string fileName = "FileDoubleCheckChange-" + today.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/PO/");
                fileName += fileExt;
                file.SaveAs(path + Path.GetFileName(fileName));
                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    for (int row = start.Row + 1; row <= end.Row; row++)
                    { // Row by row...
                        var partNo = worksheet.Cells[row, 6].Text.Trim();
                        if (partNo == null) break;
                        var dim = worksheet.Cells[row, 7].Text.Trim();
                        var drwHis = worksheet.Cells[row, 11].Text.Trim();
                        if (drwHis.Length == 1)
                        {
                            drwHis = "00" + drwHis;
                        }
                        if (drwHis.Length == 2)
                        {
                            drwHis = "0" + drwHis;
                        }
                        var ECN = worksheet.Cells[row, 13].Text.Trim();
                        var ExistPO = listPO.Where(x => x.MR.PartNo.Trim() == partNo.Trim() && x.MR.Clasification.Trim() == dim && x.MR.DrawHis == drwHis && x.MR.ECNNo == ECN).FirstOrDefault();
                        if (ExistPO == null)
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi: Sai draw his hoặc sai số ECN hoặc status PO này chưa thể double check");
                            goto exitLoop;
                        }
                        //****
                        var vendercode = worksheet.Cells[row, 4].Text.Trim();
                        if (vendercode != ExistPO.MR.OrderTo)
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi vender Code Không chính xác");
                            goto exitLoop;
                        }



                        //****
                        var deleveryKeyNo = worksheet.Cells[row, 15].Text.Trim();
                        //if (deleveryKeyNo.Length == 6)
                        //{
                        //    deleveryKeyNo = "0000" + deleveryKeyNo;
                        //}
                        //if (deleveryKeyNo.Length == 7)
                        //{
                        //    deleveryKeyNo = "000" + deleveryKeyNo;
                        //}
                        //if (deleveryKeyNo.Length == 8)
                        //{
                        //    deleveryKeyNo = "00" + deleveryKeyNo;
                        //}
                        //if (deleveryKeyNo.Length == 9)
                        //{
                        //    deleveryKeyNo = "0" + deleveryKeyNo;
                        //}
                        if (handelDeliveryKeyNo(deleveryKeyNo) != handelDeliveryKeyNo(ExistPO.Change_DeliveryKey))
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Delivery Key No Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var DeliveryDate = Convert.ToDateTime(worksheet.Cells[row, 18].Text.Trim()).ToShortDateString();
                        if (DeliveryDate != ExistPO.DeliveryDate.Value.ToShortDateString())
                        {
                            ExistPO.DeliveryDate = Convert.ToDateTime(DeliveryDate);
                            //fail.Add(partNo + "-" + dim + "Lỗi Delivery Date Không chính xác (" + DeliveryDate + ")");
                            //goto exitLoop;
                        }

                        //****
                        var deliveryDate_Change = Convert.ToDateTime(worksheet.Cells[row, 20].Text.Trim()).ToShortDateString();

                        if (ExistPO.Change_DeliveryDate != null) // Khác null thì so sánh cái mới ngược lại thì so sánh với cái current
                        {
                            if (ExistPO.Change_DeliveryDate.Value.ToShortDateString() != deliveryDate_Change)
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi Change request Delivery Date trên NPIS(" + deliveryDate_Change + ") khác với change date request trên DMS(" + ExistPO.Change_DeliveryDate.Value.ToShortDateString() + ").");
                                goto exitLoop;
                            }
                        }
                        else
                        {
                            if (ExistPO.DeliveryDate.Value.ToShortDateString() != deliveryDate_Change)
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi Change request Delivery Date trên NPIS(" + deliveryDate_Change + ") khác với change date request trên DMS(" + ExistPO.Change_DeliveryDate.Value.ToShortDateString() + ").");
                                goto exitLoop;
                            }
                        }


                        //****
                        var orderQty = worksheet.Cells[row, 24].Text.Trim();
                        if (orderQty != ExistPO.OrderQty.ToString().Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Oder Qty Không chính xác");
                            goto exitLoop;
                        }


                        //****
                        var deliveryLocation = worksheet.Cells[row, 27].Text.Trim();
                        if (deliveryLocation != ExistPO.DeliveryLocation.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Delivery Location Không chính xác");
                            goto exitLoop;
                        }


                        //****
                        var useBlockCode = worksheet.Cells[row, 29].Text.Trim();
                        if (useBlockCode != ExistPO.UseBlockCode.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Use Block Code Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var transportMethod = worksheet.Cells[row, 40].Text.Trim();
                        if (transportMethod != ExistPO.TransportMethod.Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Transport Method Không chính xác");
                            goto exitLoop;
                        }

                        //****
                        var departurePort = worksheet.Cells[row, 42].Text.Trim();
                        if (departurePort != ExistPO.TradeConditionPName.Substring(0, 5).Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Departure Port Không chính xác");
                            goto exitLoop;
                        }


                        //****
                        var arrivalPort = worksheet.Cells[row, 44].Text.Trim();
                        if (arrivalPort != ExistPO.TradeConditionPName.Substring(0, 5).Trim())
                        {
                            fail.Add(partNo + "-" + dim + "Lỗi Arrival Port Không chính xác");
                            goto exitLoop;
                        }
                        //****


                        // Xử lí khi check xong ko có lỗi
                        // Thay đổi trạng thái PO status
                        ExistPO.POStatusID = 14; //Change_W-PUR-App-NPIS
                        ExistPO.Change_PUCDoubleCheckBy = Session["Name"].ToString();
                        ExistPO.Change_PUCDoubleCheckDate = today;

                        db.Entry(ExistPO).State = EntityState.Modified;
                        db.SaveChanges();

                        // Luu vao PO Change
                        var old_PDD = ExistPO.DeliveryDate;
                        var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
                        PO_Change newPO_Change = new PO_Change();
                        newPO_Change.POID = ExistPO.POID;
                        newPO_Change.POIssueNo = ExistPO.POIssueNo;
                        newPO_Change.Old_DeliveryDate = old_PDD;
                        newPO_Change.Change_DeliveryDate = ExistPO.Change_DeliveryDate;
                        newPO_Change.Change_DeliveryKey = ExistPO.Change_DeliveryKey;
                        newPO_Change.Change_Reason = ExistPO.Change_Reason;
                        newPO_Change.Change_RequestBy = ExistPO.Change_RequestBy;
                        newPO_Change.Change_RequestDate = ExistPO.Change_RequestDate;
                        newPO_Change.Change_PURAppBy = ExistPO.Change_PURAppBy;
                        newPO_Change.Change_PURAppDate = ExistPO.Change_PURAppDate;
                        newPO_Change.Change_PUCCheckBy = ExistPO.Change_PUCCheckBy;
                        newPO_Change.Change_PUCCheckDate = ExistPO.Change_PUCCheckDate;
                        newPO_Change.Change_PUCDoubleCheckBy = ExistPO.Change_PUCDoubleCheckBy;
                        newPO_Change.Change_PUCDoubleCheckDate = ExistPO.Change_PUCDoubleCheckDate;
                        newPO_Change.Change_PURAppNPISBy = ExistPO.Change_PURAppNPISBy;
                        newPO_Change.Change_PURAppNPISDate = ExistPO.Change_PURAppNPISDate;
                        newPO_Change.Active = true;
                        newPO_Change.Evidential = ExistPO.Change_Evidential;
                        newPO_Change.ProcedureNo = currentProcedure.ProcedureNo;
                        newPO_Change.AttachmentNo = currentProcedure.Att_ChangePO;

                        db.PO_Change.Add(newPO_Change);
                        db.SaveChanges();


                        success.Add(partNo + "-" + dim + "OK");
                        suscussPONO.Add(ExistPO.POIssueNo);
                    exitLoop:
                        ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    }
                }
            }
            var data = new
            {
                suscussPONO,
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PURApproveNPISChange(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Check")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 14)
                    {


                        if (po.PaymentDate == null)
                        {
                            po.POStatusID = 8; // W-Payment
                        }
                        else
                        {
                            po.POStatusID = 23; // Paid
                        }
                        po.Change_PURAppNPISBy = Session["Name"].ToString();
                        po.Change_PURAppNPISDate = DateTime.Now;

                        // gán lại Delivery date sau khi được change mới

                        po.DeliveryDate = po.Change_DeliveryDate;

                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult purRejectChange(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 11)
                    {
                        po.POStatusID = 15; // Change_W-PUR-CONFIRM
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pucRejectChange(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For reject
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Approve" || role == "Check"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 12 || po.POStatusID == 13)
                    {
                        po.POStatusID = 15; // Back_Change_W-PUR-CONFIRM
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult importChangePOFile(string id, HttpPostedFileBase file) // Change PO bằng file excell
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            List<string> suscussPONO = new List<string>();
            if (dept == "PUR" && (role == "Check" || role == "Approve"))
            {
                // Luu file Execl
                string fileName = "FileImportChangePO-" + today.ToString("yyyy-MM-dd-hhss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/PO/");
                fileName += fileExt;
                file.SaveAs(path + Path.GetFileName(fileName));
                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    for (int row = start.Row + 13; row <= end.Row; row++)
                    { // Row by row...
                        var venderCode = worksheet.Cells[row, 2].Text.Trim().ToUpper();
                        var partNo = worksheet.Cells[row, 3].Text.Trim().ToUpper();
                        var dim = worksheet.Cells[row, 4].Text.Trim().ToUpper();

                        var his = worksheet.Cells[row, 14].Text.Trim().ToUpper();
                        var ECN = worksheet.Cells[row, 16].Text.Trim().ToUpper();

                        if (String.IsNullOrEmpty(partNo)) break;
                        var ExistPO = db.PO_Dies.Where(x => x.MR.PartNo == partNo && x.MR.Clasification == dim
                        && x.MR.DrawHis == his && x.MR.ECNNo == ECN && (x.POStatusID == 8 || x.POStatusID == 15)).FirstOrDefault();
                        if (ExistPO == null)
                        {
                            fail.Add(partNo + "-" + dim + " Lỗi Không tìm thấy hoặc Change PO này. Note: Chỉ có thể Change PO nếu status là: W-PAYMENT, Back_Change_W-PUR-CONFIRM");
                            goto exitLoop;
                        }
                        //****
                        var change_DeliveryDate = worksheet.Cells[row, 7].Text.Trim();
                        if (!String.IsNullOrEmpty(change_DeliveryDate))
                        {
                            try
                            {
                                var converDateDelivery = Convert.ToDateTime(change_DeliveryDate);
                                ExistPO.Change_DeliveryDate = converDateDelivery;
                            }
                            catch
                            {
                                fail.Add(partNo + "-" + dim + " Lỗi Delivery Date ko đúng định dạng ngày!");
                                goto exitLoop;
                            }
                        }
                        //****
                        var deliveryKeyNo = worksheet.Cells[row, 5].Text.Trim().ToUpper();
                        if (!String.IsNullOrEmpty(deliveryKeyNo))
                        {
                            try
                            {
                                if (deliveryKeyNo.Length > 7)
                                {
                                    ExistPO.Change_DeliveryKey = deliveryKeyNo;
                                }
                                else
                                {
                                    fail.Add(partNo + "-" + dim + " Lỗi Delivery No phải có ít nhất 8 kí tự!");
                                    goto exitLoop;
                                }
                            }
                            catch
                            {
                                fail.Add(partNo + "-" + dim + " Lỗi Delivery Date ko đúng định dạng ngày!");
                                goto exitLoop;
                            }
                        }



                        // Xử lí khi check xong ko có lỗi
                        ExistPO.POStatusID = 11; // Change_W-M1-PUR- App
                        ExistPO.Change_Reason = "Change PO Die To Make Payment";

                        // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                        var oldPOIssueNo = ExistPO.POIssueNo;
                        // Tạo số PO mới
                        ExistPO.POIssueNo = createPOIssueNo(ExistPO.POIssueNo, false, true, false, false);
                        ExistPO.Change_RequestBy = Session["Name"].ToString();
                        ExistPO.Change_RequestDate = today;
                        // Luu DB
                        db.Entry(ExistPO).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(ExistPO.POIssueNo);
                    exitLoop:
                        ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        // *************KẾT THÚC CHANGE PO


        // ************Cancel PO Area
        // Nguyên tắc: 
        // 1. Tất cả PO có thể cancel
        // 2. Sẽ ghi nhập PO Cancel sau khi PUR App NPIS => POStatus = 14 Change_W-PUR-APP NPIS
        // 3. Hủy các kết quả create die
        // 4. Copy new PO to reissue (Incase Reissue = true)
        public JsonResult checkAllowCancel(string[] id)
        {
            var status = "";
            var po = db.PO_Dies.Find(Convert.ToInt32(id[0]));
            if (po.POstatusCalogory.isCanCancelPOWithoutRoute == true) // W-issue / W-PUR-M1(GM)-App /  W-PUC-INput NPIS / Back
            {
                status = "NoRoute";
            }
            else
            {
                if (po.POstatusCalogory.isCanCancelPOWithoutRoute == false) // W-DbCheck / W-PUR App NPIS / W-Price/ W-App-Price / W-Pay
                {
                    status = "Route";
                }
                else
                {
                    status = "NG";
                }
            }

            return Json(status, JsonRequestBehavior.AllowGet);
        }
        public JsonResult cancelPO(string[] id, string reason, string deliveryKeyNo, HttpPostedFileBase file)
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();

            // for Confirm 

            if (dept == "PUR")
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[0]));
                // nếu có file => Luu file
                if (file != null) // Luu file
                {
                    string fileName = "Evidential-" + today.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/PO/Evidential/");
                    fileName += fileExt;
                    file.SaveAs(path + Path.GetFileName(fileName));
                    po.Cancel_Evidential = fileName;
                }
                // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                var oldPOIssueNo = po.POIssueNo;
                // Tạo số PO mới
                po.POIssueNo = createPOIssueNo(po.POIssueNo, false, false, true, false);
                po.ProcedureNo = currentProcedure.ProcedureNo;
                po.AttachNo_Cancel = currentProcedure.Att_CancelPO;
                if (po.POstatusCalogory.isCanCancelPOWithoutRoute == false) // Fải chạy route qua PUC
                {
                    if (!String.IsNullOrEmpty(reason) && !String.IsNullOrEmpty(deliveryKeyNo))
                    {
                        po.POStatusID = 16; // Cancel_W-PUR-M1-APPROVE
                        po.Cancel_Reason = reason;
                        //if (!String.IsNullOrEmpty(deliveryKeyNo))
                        //{
                        //    if (deliveryKeyNo.Length == 6)
                        //    {
                        //        deliveryKeyNo = deliveryKeyNo + "0000";
                        //    }
                        //    if (deliveryKeyNo.Length == 7)
                        //    {
                        //        deliveryKeyNo = deliveryKeyNo + "000";
                        //    }
                        //    if (deliveryKeyNo.Length == 8)
                        //    {
                        //        deliveryKeyNo = deliveryKeyNo + "00";
                        //    }
                        //    if (deliveryKeyNo.Length == 9)
                        //    {
                        //        deliveryKeyNo = deliveryKeyNo + "0";
                        //    }
                        //}
                        po.Cancel_DeliveryKey = deliveryKeyNo;
                        po.Reissue = false;
                        po.Cancel_RequestBy = Session["Name"].ToString();
                        po.Cancel_RequestDate = today;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.MR.MRNo);
                    }
                    else
                    {
                        fail.Add(po.MR.MRNo);
                    }
                }
                if (po.POstatusCalogory.isCanCancelPOWithoutRoute == true)
                {
                    //1. Cancel PO => Ko cần App
                    po.POStatusID = 20; // cancel
                    po.Cancel_Reason = reason;
                    po.Cancel_DeliveryKey = null;
                    db.Entry(po).State = EntityState.Modified;
                    //2. Cancel MR =>
                    var mR = db.MRs.Find(po.MRID);
                    mR.StatusID = 12; // cancel
                    mR.Note = DateTime.Now.ToString("yyyy-MM-dd: ") + Session["Name"] + " Cancel MR & PO Reason: " + reason + System.Environment.NewLine + mR.Note;
                    db.Entry(mR).State = EntityState.Modified;
                    db.SaveChanges();
                    success.Add(po.MR.MRNo);
                    cancelDie_Part_Common(po.MRID);
                }

            }
            var data = new
            {
                success,
                fail
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult cancelPO_reIssue(string[] id, string reason, string deliveryKeyNo, HttpPostedFileBase file)
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();

            // for Confirm 

            if (dept == "PUR")
            {
                var po = db.PO_Dies.Find(Convert.ToInt32(id[0]));
                // nếu có file => Luu file
                if (file != null) // Luu file
                {
                    string fileName = "Evidential-" + today.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/PO/Evidential/");
                    fileName += fileExt;
                    file.SaveAs(path + Path.GetFileName(fileName));
                    po.Cancel_Evidential = fileName;
                }
                // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                var oldPOIssueNo = po.POIssueNo;
                // Tạo số PO mới
                po.POIssueNo = createPOIssueNo(po.POIssueNo, false, false, true, false);
                po.ProcedureNo = currentProcedure.ProcedureNo;
                po.AttachNo_Cancel = currentProcedure.Att_CancelPO;
                if (po.POstatusCalogory.isCanCancelPOWithoutRoute == false) // Fải chạy route qua PUC
                {
                    if (!String.IsNullOrEmpty(reason) && !String.IsNullOrEmpty(deliveryKeyNo))
                    {
                        po.POStatusID = 16; // Cancel_W-PUR-M1-APPROVE
                        po.Cancel_Reason = reason;

                        po.Cancel_DeliveryKey = deliveryKeyNo;
                        po.Reissue = true;
                        po.Cancel_RequestBy = Session["Name"].ToString();
                        po.Cancel_RequestDate = today;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.MR.MRNo);
                    }
                    else
                    {
                        fail.Add(po.MR.MRNo);
                    }
                }
                if (po.POstatusCalogory.isCanCancelPOWithoutRoute == true)
                {
                    //1. Cancel PO => Ko cần App
                    po.POStatusID = 20; // cancel
                    po.Cancel_Reason = reason;
                    po.Cancel_DeliveryKey = null;
                    db.Entry(po).State = EntityState.Modified;


                    //2. Reissue => Copy New PO
                    PO_Dies newPOReissue = new PO_Dies();
                    if (po.MR.EstimateCost == 10)
                    {
                        newPOReissue.TempPO = true;
                    }
                    else
                    {
                        newPOReissue.TempPO = false;
                    }
                    var PR = po.MR.Clasification.ToString();
                    newPOReissue.POIssueNo = po.POIssueNo;
                    newPOReissue.PR = PR.Remove(PR.Length - 1);
                    newPOReissue.MRID = po.MR.MRID;
                    newPOReissue.POStatusID = 21; // W-PUR-Re-ISSUE
                    newPOReissue.Active = true;
                    newPOReissue.DeliveryDate = po.MR.PDD;
                    newPOReissue.CreateDate = today;
                    // 
                    newPOReissue.OriginalDeliveryDate = po.OriginalDeliveryDate;
                    newPOReissue.PODate = po.PODate;
                    newPOReissue.OriginalPOdate = po.OriginalPOdate;
                    newPOReissue.Price = po.Price;
                    newPOReissue.PUSCheckBy = po.PUSCheckBy;
                    newPOReissue.PUSCheckDate = po.PUSCheckDate;
                    newPOReissue.PUSAppBy = po.PUSAppBy;
                    newPOReissue.PUSAppDate = po.PUSAppDate;

                    db.PO_Dies.Add(newPOReissue);
                    db.SaveChanges();
                    success.Add(po.MR.MRNo);
                }

            }
            var data = new
            {
                success,
                fail
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PurAPPCancel(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1 || GM
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 16) // cancel_W-PUR-M1-App
                    {
                        po.POStatusID = 17; // cancel_W-PUC-Input-NPIS
                        po.Cancel_PURAppBy = Session["Name"].ToString();
                        po.Cancel_PURAppDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }

            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pucConfirmCancel(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Check" || role == "Approve"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 17) //cancel_W-PUS-double Check
                    {
                        po.POStatusID = 18; // cancel_W-PUS-double Check
                        po.Cancel_PUCCheckBy = Session["Name"].ToString();
                        po.Cancel_PUCCheckDate = DateTime.Now;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult doubleCheckCancel(string id, HttpPostedFileBase file, bool isHandCheck)
        {
            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            List<string> suscussPONO = new List<string>();
            if (dept == "PUC" && role == "Approve")
            {
                var listPO = db.PO_Dies.Where(x => x.POStatusID == 18 && x.Active != false).ToList();
                if (file != null && isHandCheck == false)
                {
                    // Luu file Execl
                    string fileName = "FileDoubleCheckCancel-" + today.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/PO/");
                    fileName += fileExt;
                    file.SaveAs(path + Path.GetFileName(fileName));
                    // Doc file
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                        var start = worksheet.Dimension.Start;
                        var end = worksheet.Dimension.End;

                        for (int row = start.Row + 1; row <= end.Row; row++)
                        { // Row by row...
                            var partNo = worksheet.Cells[row, 6].Text.Trim();

                            if (partNo == null) break;
                            var dim = worksheet.Cells[row, 7].Text.Trim();
                            var drwHis = worksheet.Cells[row, 11].Text.Trim();
                            if (drwHis.Length == 1)
                            {
                                drwHis = "00" + drwHis;
                            }
                            if (drwHis.Length == 2)
                            {
                                drwHis = "0" + drwHis;
                            }
                            var ECN = worksheet.Cells[row, 12].Text.Trim();
                            var ExistPO = listPO.Where(x => x.MR.PartNo.Trim() == partNo.Trim() && x.MR.Clasification.Trim() == dim.Trim() && x.MR.DrawHis == drwHis && x.MR.ECNNo == ECN).FirstOrDefault();
                            if (ExistPO == null)
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi: Sai draw his hoặc sai số ECN hoặc status PO này chưa thể double check");
                                goto exitLoop;
                            }
                            //****
                            var vendercode = worksheet.Cells[row, 4].Text.Trim();
                            if (vendercode != ExistPO.MR.OrderTo)
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi vender Code Không chính xác");
                                goto exitLoop;
                            }
                            //****
                            var deleveryKeyNo = worksheet.Cells[row, 16].Text.Trim();


                            if (handelDeliveryKeyNo(deleveryKeyNo) != handelDeliveryKeyNo(ExistPO.Cancel_DeliveryKey))
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi Delivery Key No Không chính xác");
                                goto exitLoop;
                            }

                            //****
                            var DeliveryDate = Convert.ToDateTime(worksheet.Cells[row, 20].Text.Trim()).ToShortDateString();
                            if (DeliveryDate != ExistPO.DeliveryDate.Value.ToShortDateString())
                            {
                                fail.Add(partNo + "-" + dim + "Lỗi Delivery Date Không chính xác");
                                goto exitLoop;
                            }

                            // Xử lí khi check xong ko có lỗi
                            // Thay đổi trạng thái PO status
                            ExistPO.POStatusID = 19; //W-PUR-App-NPIS

                            ExistPO.Cancel_PUCDoubleCheckBy = Session["Name"].ToString();
                            ExistPO.Cancel_PUCDoubleCheckDate = today;
                            db.Entry(ExistPO).State = EntityState.Modified;
                            db.SaveChanges();

                            // Luu vao PO Cancel
                            var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
                            PO_Cancel newPO_Cancel = new PO_Cancel();
                            newPO_Cancel.POID = ExistPO.POID;
                            newPO_Cancel.POIssueNo = ExistPO.POIssueNo;
                            newPO_Cancel.Cancel_DeliveryKey = ExistPO.Cancel_DeliveryKey;
                            newPO_Cancel.Cancel_Reason = ExistPO.Cancel_Reason;
                            newPO_Cancel.Cancel_RequestBy = ExistPO.Cancel_RequestBy;
                            newPO_Cancel.Cancel_RequestDate = ExistPO.Cancel_RequestDate;
                            newPO_Cancel.Cancel_PURAppBy = ExistPO.Cancel_PURAppBy;
                            newPO_Cancel.Cancel_PURAppDate = ExistPO.Cancel_PURAppDate;
                            newPO_Cancel.Cancel_PUCCheckBy = ExistPO.Cancel_PUCCheckBy;
                            newPO_Cancel.Cancel_PUCCheckDate = ExistPO.Cancel_PUCCheckDate;
                            newPO_Cancel.Cancel_PUCDouleCheckBy = ExistPO.Cancel_PUCDoubleCheckBy;
                            newPO_Cancel.Cancel_PUCDoubleCheckDate = ExistPO.Cancel_PUCDoubleCheckDate;
                            newPO_Cancel.Cancel_PURAppNPISBy = ExistPO.Cancel_PURAppNPISBy;
                            newPO_Cancel.Cancel_PURAppNPISDate = ExistPO.Cancel_PURAppNPISDate;
                            newPO_Cancel.Active = true;
                            newPO_Cancel.ProcedureNo = currentProcedure.ProcedureNo;
                            newPO_Cancel.AttachmentNo = currentProcedure.Att_CancelPO;
                            newPO_Cancel.Evidential = ExistPO.Cancel_Evidential;
                            db.PO_Cancel.Add(newPO_Cancel);
                            db.SaveChanges();

                            success.Add(partNo + "-" + dim + "OK");
                            suscussPONO.Add(ExistPO.POIssueNo);
                        exitLoop:
                            ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                        }
                    }
                }
                else
                {
                    var ExistPO = db.PO_Dies.Find(int.Parse(id));
                    // Xử lí khi check xong ko có lỗi
                    // Thay đổi trạng thái PO status
                    ExistPO.POStatusID = 19; //W-PUR-App-NPIS

                    ExistPO.Cancel_PUCDoubleCheckBy = Session["Name"].ToString();
                    ExistPO.Cancel_PUCDoubleCheckDate = today;
                    db.Entry(ExistPO).State = EntityState.Modified;
                    db.SaveChanges();

                    // Luu vao PO Cancel
                    var currentProcedure = db.POProcedures.Where(x => x.Type.Contains("PO") && x.EffectiveDate < DateTime.Now && x.Active != false).OrderByDescending(x => x.EffectiveDate).FirstOrDefault();
                    PO_Cancel newPO_Cancel = new PO_Cancel();
                    newPO_Cancel.POID = ExistPO.POID;
                    newPO_Cancel.POIssueNo = ExistPO.POIssueNo;
                    newPO_Cancel.Cancel_DeliveryKey = ExistPO.Cancel_DeliveryKey;
                    newPO_Cancel.Cancel_Reason = ExistPO.Cancel_Reason;
                    newPO_Cancel.Cancel_RequestBy = ExistPO.Cancel_RequestBy;
                    newPO_Cancel.Cancel_RequestDate = ExistPO.Cancel_RequestDate;
                    newPO_Cancel.Cancel_PURAppBy = ExistPO.Cancel_PURAppBy;
                    newPO_Cancel.Cancel_PURAppDate = ExistPO.Cancel_PURAppDate;
                    newPO_Cancel.Cancel_PUCCheckBy = ExistPO.Cancel_PUCCheckBy;
                    newPO_Cancel.Cancel_PUCCheckDate = ExistPO.Cancel_PUCCheckDate;
                    newPO_Cancel.Cancel_PUCDouleCheckBy = ExistPO.Cancel_PUCDoubleCheckBy;
                    newPO_Cancel.Cancel_PUCDoubleCheckDate = ExistPO.Cancel_PUCDoubleCheckDate;
                    newPO_Cancel.Cancel_PURAppNPISBy = ExistPO.Cancel_PURAppNPISBy;
                    newPO_Cancel.Cancel_PURAppNPISDate = ExistPO.Cancel_PURAppNPISDate;
                    newPO_Cancel.Active = true;
                    newPO_Cancel.ProcedureNo = currentProcedure.ProcedureNo;
                    newPO_Cancel.AttachmentNo = currentProcedure.Att_CancelPO;
                    newPO_Cancel.Evidential = ExistPO.Cancel_Evidential;
                    db.PO_Cancel.Add(newPO_Cancel);
                    db.SaveChanges();

                    success.Add(ExistPO.MR.PartNo + "-" + ExistPO.MR.Clasification + " - OK");
                    suscussPONO.Add(ExistPO.POIssueNo);
                }
            }
            var data = new
            {
                suscussPONO,
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult PURApproveNPISCancel(string[] id)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();
            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // for Confirm 
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Check")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 19)
                    {
                        po.POStatusID = 20; // CANCELLED
                        po.Cancel_PURAppNPISBy = Session["Name"].ToString();
                        po.Cancel_PURAppNPISDate = DateTime.Now;


                        db.Entry(po).State = EntityState.Modified;

                        if (po.Reissue != true) // CancelPO =>Cancel PO
                        {
                            var mR = db.MRs.Find(po.MRID);
                            mR.StatusID = 12;// Canceled MR
                            mR.Note = DateTime.Now.ToString("yyyy-MM-dd: ") + Session["Name"] + " Cancel MR & PO Reason: " + po.Cancel_Reason + System.Environment.NewLine + mR.Note;
                            db.Entry(mR).State = EntityState.Modified;
                            db.SaveChanges();
                            //Huy cac ket noi to die da duoc tao ra
                            cancelDie_Part_Common(po.MRID);
                        }
                        else // CancelPO & Re-issue
                        {
                            //2. Reissue => Copy New PO
                            PO_Dies newPOReissue = new PO_Dies();
                            if (po.MR.EstimateCost == 10)
                            {
                                newPOReissue.TempPO = true;
                            }
                            else
                            {
                                newPOReissue.TempPO = false;
                            }
                            var PR = po.MR.Clasification.ToString();
                            newPOReissue.POIssueNo = po.POIssueNo;
                            newPOReissue.PR = PR.Remove(PR.Length - 1);
                            newPOReissue.MRID = po.MR.MRID;
                            newPOReissue.POStatusID = 21; // W-PUR-Re-ISSUE
                            newPOReissue.Active = true;

                            newPOReissue.CreateDate = DateTime.Now;
                            // newPOReissue.OriginalPOdate = po.OriginalPOdate != null ? po.OriginalPOdate : po.PODate;
                            newPOReissue.DeliveryDate = po.DeliveryDate;
                            newPOReissue.OriginalDeliveryDate = po.OriginalDeliveryDate;
                            newPOReissue.PODate = po.PODate;
                            newPOReissue.OriginalPOdate = po.OriginalPOdate;
                            newPOReissue.Price = po.Price;
                            newPOReissue.PUSCheckBy = po.PUSCheckBy;
                            newPOReissue.PUSCheckDate = po.PUSCheckDate;
                            newPOReissue.PUSAppBy = po.PUSAppBy;
                            newPOReissue.PUSAppDate = po.PUSAppDate;
                            db.PO_Dies.Add(newPOReissue);
                            db.SaveChanges();
                        }

                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult purRejectCancel(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For Chi Giang PUR M1
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUR" && role == "Approve")
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 16)
                    {
                        po.POStatusID = 22; // Change_W-PUR-CONFIRM
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd  ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }

            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public JsonResult pucRejectCancel(string[] id, string reason)
        {
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["PO_Role"].ToString().Trim();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            // For reject
            for (var i = 0; i < id.Length; i++)
            {
                if (dept == "PUC" && (role == "Approve" || role == "Check"))
                {
                    var po = db.PO_Dies.Find(Convert.ToInt32(id[i]));
                    if (po.POStatusID == 17 || po.POStatusID == 18)
                    {
                        po.POStatusID = 22; // Back_Cancel_W-PUR-CONFIRM
                        po.Remark = DateTime.Now.ToString("yyyy/MM/dd  ") + Session["Name"].ToString() + " Reject :" + reason + System.Environment.NewLine + po.Remark;
                        db.Entry(po).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(po.POIssueNo);
                    }
                    else
                    {
                        fail.Add(po.POIssueNo);
                    }
                }
            }
            var data = new
            {
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }

        // *************KẾT THÚC Cancel PO






        // Common
        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-excel", fileName);
            }
            else
            {
                // Problem - Log the error, generate a blank file,
                //           redirect to another controller action - whatever fits with your application
                return new EmptyResult();
            }
        }
        public ActionResult exportPOToControlList(List<PO_Dies> poes)
        {
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/Format/PO_Control_List02.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["PO_Control"];
                int rowId = 6;
                int i = 1;

                foreach (var po in poes)
                {
                    sheet.Cells["A" + rowId.ToString()].Value = i;
                    sheet.Cells["B" + rowId.ToString()].Value = po.POstatusCalogory.Status;
                    sheet.Cells["C" + rowId.ToString()].Value = po.MR.MRNo;
                    sheet.Cells["D" + rowId.ToString()].Value = po.POIssueNo;
                    sheet.Cells["E" + rowId.ToString()].Value = po.MR.DieNo;
                    sheet.Cells["F" + rowId.ToString()].Value = po.MR.PartNo;
                    sheet.Cells["G" + rowId.ToString()].Value = po.MR.Clasification;
                    sheet.Cells["H" + rowId.ToString()].Value = po.MR.MRType.Type;
                    sheet.Cells["I" + rowId.ToString()].Value = po.MR.PartName;
                    sheet.Cells["J" + rowId.ToString()].Value = po.TempPO == true ? "Y" : "N";
                    sheet.Cells["K" + rowId.ToString()].Value = po.ReasonIssuePO;
                    sheet.Cells["L" + rowId.ToString()].Value = po.POstatusCalogory.Status.ToUpper().Contains("CANCEL_") ? po.Cancel_DeliveryKey : po.Change_DeliveryKey;
                    sheet.Cells["M" + rowId.ToString()].Value = po.MR.ModelName;
                    sheet.Cells["N" + rowId.ToString()].Value = po.MR.Supplier.SupplierCode;
                    sheet.Cells["O" + rowId.ToString()].Value = po.MR.Supplier.SupplierName;
                    sheet.Cells["P" + rowId.ToString()].Value = po.MR.OrderTo;
                    sheet.Cells["Q" + rowId.ToString()].Value = po.VendorFctry;
                    sheet.Cells["R" + rowId.ToString()].Value = po.NeedRegisterRateTableByPart;
                    sheet.Cells["S" + rowId.ToString()].Value = po.MR.ProcessCodeCalogory.Type;
                    sheet.Cells["T" + rowId.ToString()].Value = po.MR.DrawHis;
                    sheet.Cells["U" + rowId.ToString()].Value = po.MR.ECNNo;
                    sheet.Cells["V" + rowId.ToString()].Value = po.DeliveryDate;
                    sheet.Cells["W" + rowId.ToString()].Value = po.OrderQty;
                    sheet.Cells["X" + rowId.ToString()].Value = po.DeliveryLocation;
                    sheet.Cells["Y" + rowId.ToString()].Value = po.UseBlockCode;
                    sheet.Cells["Z" + rowId.ToString()].Value = po.ItemCategory;
                    sheet.Cells["AA" + rowId.ToString()].Value = po.TransportMethod;
                    sheet.Cells["AB" + rowId.ToString()].Value = po.ContainerLoadingCode;
                    sheet.Cells["AC" + rowId.ToString()].Value = po.MR.CavQty;
                    sheet.Cells["AD" + rowId.ToString()].Value = po.WarrantyShot;
                    sheet.Cells["AE" + rowId.ToString()].Value = po.TradeConditionPName;
                    sheet.Cells["AF" + rowId.ToString()].Value = po.MR.ModelName;
                    sheet.Cells["AG" + rowId.ToString()].Value = po.IssueBy;
                    sheet.Cells["AH" + rowId.ToString()].Value = po.IssueDate.HasValue ? po.IssueDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AI" + rowId.ToString()].Value = po.PURAppBy;
                    sheet.Cells["AJ" + rowId.ToString()].Value = po.PURAppDate.HasValue ? po.PURAppDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AK" + rowId.ToString()].Value = po.GMPURAppBy;
                    sheet.Cells["AL" + rowId.ToString()].Value = po.GMPURAppDate.HasValue ? po.GMPURAppDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AM" + rowId.ToString()].Value = po.PUCCheckBy;
                    sheet.Cells["AN" + rowId.ToString()].Value = po.PUCCheckDate.HasValue ? po.PUCCheckDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AO" + rowId.ToString()].Value = po.PUCDoubleCheckBy;
                    sheet.Cells["AP" + rowId.ToString()].Value = po.PUCDoubleCheckDate.HasValue ? po.PUCDoubleCheckDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AQ" + rowId.ToString()].Value = po.PURAppNPISBy;
                    sheet.Cells["AR" + rowId.ToString()].Value = po.PURAppNPISDate.HasValue ? po.PURAppNPISDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AS" + rowId.ToString()].Value = po.PODate;
                    sheet.Cells["AT" + rowId.ToString()].Value = po.Price;
                    sheet.Cells["AU" + rowId.ToString()].Value = po.MR.Unit;
                    sheet.Cells["AV" + rowId.ToString()].Value = po.Remark;
                    sheet.Cells["AW" + rowId.ToString()].Value = po.PUSCheckDate.HasValue ? po.PUSCheckDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["AX" + rowId.ToString()].Value = "";
                    sheet.Cells["AY" + rowId.ToString()].Value = po.PUSCheckBy;
                    sheet.Cells["AZ" + rowId.ToString()].Value = po.PUSAppBy;
                    sheet.Cells["BA" + rowId.ToString()].Value = po.PUSAppDate.HasValue ? po.PUSAppDate.Value.ToString("yyyy-MM-dd") : "-";
                    sheet.Cells["BB" + rowId.ToString()].Value = po.PaymentDate.HasValue ? po.PaymentDate.Value.ToString("yyyy-MM-dd") : "-";

                    i++;
                    rowId++;
                }

                package.SaveAs(output);
            }
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=PO_Control_List" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx");
            output.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();
            return RedirectToAction("Index");
        }




        public JsonResult importPaymentDateAndPrice(HttpPostedFileBase file, string action)
        {
            var msg = "";
            int success = 0;
            int fail = 0;
            var dept = Session["Dept"].ToString();
            var today = DateTime.Now;
            // Luu file Execl
            string fileName = "Import_Payment_Or_Price-" + today.ToString("yyyy-MM-dd-hhmmss");
            string fileExt = Path.GetExtension(file.FileName);
            string path = Server.MapPath("~/File/PO/");
            fileName += fileExt;
            file.SaveAs(path + Path.GetFileName(fileName));

            if (action == "Import Payment" && dept == "PUR")
            {
                // var listPO = db.PO_Dies.Where(x => x.POStatusID == 8 && x.Active != false).ToList();

                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    var rowVender = 0;
                    var rowPartNo = 0;
                    var rowDim = 0;
                    var rowDraw = 0;
                    var rowECN = 0;
                    var rowPayDate = 0;

                    // Tìm cột nào là Parts No + Dim + ECN_No
                    for (int i = 1; i <= end.Column; i++)
                    {
                        if (worksheet.Cells[1, i].Text == "CD_SPLY") // Maker
                        {
                            rowVender = i;
                        }
                        if (worksheet.Cells[1, i].Text == "NO_PARTS") // part No
                        {
                            rowPartNo = i;
                        }
                        if (worksheet.Cells[1, i].Text == "NO_ADJ_DIM") // Dim
                        {
                            rowDim = i;
                        }
                        if (worksheet.Cells[1, i].Text == "NO_DRAW") // his
                        {
                            rowDraw = i;
                        }
                        if (worksheet.Cells[1, i].Text == "CD_CHG_HIST_ALL") // ECN
                        {
                            rowECN = i;
                        }
                        if (worksheet.Cells[1, i].Text == "DT_REC") // Payment Date
                        {
                            rowPayDate = i;
                        }
                    }



                    for (int row = start.Row + 1; row <= end.Row; row++)
                    { // Row by row...
                        var vender = worksheet.Cells[row, rowVender].Text.Trim();
                        var partNo = worksheet.Cells[row, rowPartNo].Text.Trim();
                        var dim = worksheet.Cells[row, rowDim].Text.Trim();
                        var his = worksheet.Cells[row, rowDraw].Text.Trim();
                        var ecn = worksheet.Cells[row, rowECN].Text.Trim();
                        var paymentDate = worksheet.Cells[row, rowPayDate].Text.Trim();
                        // find PO exist ?
                        var po = db.PO_Dies.Where(x => x.MR.OrderTo == vender && x.MR.PartNo == partNo && x.MR.Clasification == dim && x.MR.DrawHis == his && x.MR.ECNNo == ecn && x.Active != false && x.POStatusID != 20).FirstOrDefault();
                        if (po != null)
                        {
                            var newDatepayment = DateTime.ParseExact(paymentDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                            po.PaymentDate = newDatepayment;
                            if (po.POStatusID == 8)
                            {
                                po.POStatusID = 23; // paid
                            }
                            var resultExchange = commonFunction.exchangeToUSD(po.Price, po.MR.Unit);
                            db.Entry(po).State = EntityState.Modified;

                            // Update DieStatus => W-MP
                            if (po.MR.TypeID < 4) // nhỏ hơn 4 là new/add/Renew
                            {
                                var die = db.Die1.Where(x => x.DieNo == po.MR.DieNo && x.Active != false && x.isCancel != true && x.isOfficial == true).FirstOrDefault();
                                if (die != null)
                                {
                                    die.DieStatusID = 2; // W-MP
                                    db.Entry(die).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }



                            // Update MR
                            var mR = db.MRs.Find(po.MRID);
                            mR.ExchangeRate = resultExchange.rate;
                            mR.AppCostExchangeUSD = resultExchange.price;
                            mR.StatusID = 14;
                            db.Entry(mR).State = EntityState.Modified;
                            db.SaveChanges();
                            success++;

                        }
                        else
                        {
                            fail++;
                            goto exitLoop;
                        }
                    exitLoop:
                        ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    }
                    msg = "OK";

                }
            }
            if (action == "Import Price" && dept == "PUS")
            {
                var listPO = db.PO_Dies.Where(x => (x.POStatusID == 6 || x.POStatusID == 7 || x.POStatusID == 8 || x.POStatusID == 24) && x.Active != false).ToList();

                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileName))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    var rowVender = 0;
                    var rowPartNo = 0;
                    var rowDim = 0;
                    var rowDraw = 0;
                    var rowECN = 0;
                    var rowPrice = 0;
                    var rowPoDate = 0;

                    // Tìm cột nào là Parts No + Dim + ECN_No
                    for (int i = 1; i <= end.Column; i++)
                    {
                        if (worksheet.Cells[1, i].Text == "Vendor")
                        {
                            rowVender = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Parts No")
                        {
                            rowPartNo = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Dim")
                        {
                            rowDim = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Drawing No.")
                        {
                            rowDraw = i;
                        }
                        if (worksheet.Cells[1, i].Text == "ECN No")
                        {
                            rowECN = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Purchase Price")
                        {
                            rowPrice = i;
                        }
                        if (worksheet.Cells[1, i].Text == "UP Reg Dt")
                        {
                            rowPoDate = i;
                        }
                    }


                    for (int row = start.Row + 1; row <= end.Row; row++)
                    { // Row by row...
                        var vender = worksheet.Cells[row, rowVender].Text.Trim();
                        var partNo = worksheet.Cells[row, rowPartNo].Text.Trim();
                        var dim = worksheet.Cells[row, rowDim].Text.Trim();
                        var ECNno = worksheet.Cells[row, rowECN].Text.Trim();
                        var his = worksheet.Cells[row, rowDraw].Text.Trim();
                        if (his.Length == 1)
                        {
                            his = "00" + his;
                        }
                        if (his.Length == 2)
                        {
                            his = "0" + his;
                        }
                        var price = worksheet.Cells[row, rowPrice].Text.Trim();
                        var poDate = worksheet.Cells[row, rowPoDate].Text.Trim();

                        double intPrice;
                        bool isPrice = double.TryParse(price, out intPrice);
                        DateTime Date_poDate;
                        bool isPoDate = DateTime.TryParse(poDate, out Date_poDate);
                        if (isPrice == false)
                        {
                            goto exitLoop;
                        }
                        if (isPoDate == false)
                        {
                            goto exitLoop;
                        }

                        // find PO exist ?
                        var po = listPO.Where(x => x.MR.PartNo == partNo && x.MR.Clasification == dim && x.MR.DrawHis == his && x.MR.ECNNo == ECNno && x.MR.OrderTo == vender && x.Active != false).FirstOrDefault();
                        if (po != null)
                        {
                            po.Price = intPrice;
                            po.PODate = Date_poDate;
                            po.OriginalPOdate = po.OriginalPOdate != null ? po.OriginalPOdate : po.PODate;
                            po.POStatusID = 7;
                            po.PUSCheckBy = Session["Name"].ToString();
                            po.PUSCheckDate = today;
                            db.Entry(po).State = EntityState.Modified;
                            db.SaveChanges();
                            commonFunction.genarateNewDie("", "", "", "", "", "", "", "", "", po.MR, po.WarrantyShot);
                            if (po.TempPO == true && po.Price == 10)
                            {
                                sendMailJob.sendEmailTempPO("PUSAPP", Session["Mail"].ToString(), po.MR.PartNo, po.MR.Clasification, po.MR.OrderTo, "Need PUS Approval Temporary Price.");
                            }
                            success++;
                            msg = msg + partNo + "-OK";
                        }
                        else
                        {
                            fail++;
                            msg = msg + partNo + "-NG";
                            goto exitLoop;
                        }
                    exitLoop:
                        ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    }

                }

            }
            var data = new
            {
                msg = msg,
                success,
                fail
            };

            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public string createPOIssueNo(string poIssueNo, bool isNew, bool isChange, bool isCancel, bool isReIsuse)
        {
            var today = DateTime.Now;
            string SucPOIssueNo = "";

            var totalPOinYear = db.PO_Dies.Count(x => x.POIssueNo != null && x.IssueDate.Value.Year == today.Year) + 1;
            // Neu new Issue
            if (isNew) // mới mà chưa có số PO
            {
                if (String.IsNullOrEmpty(poIssueNo))
                {
                    SucPOIssueNo = "PO" + today.ToString("yyMMdd") + "-" + totalPOinYear + "-NE-00";
                }
                else
                {
                    SucPOIssueNo = poIssueNo;
                }
            }
            else
            {
                var mainNo = "";
                var upverStr = "";
                if (!String.IsNullOrEmpty(poIssueNo))
                {
                    var curentPOissueNo = poIssueNo;
                    mainNo = curentPOissueNo.Remove(curentPOissueNo.Length - 5, 5); //NE-00
                    var upver = curentPOissueNo.Substring(curentPOissueNo.Length - 2, 2); //00
                    int upverInt = Convert.ToInt16(upver) + 1;
                    upverStr = Convert.ToString(upverInt);
                    if (upverStr.Length == 1)
                    {
                        upverStr = "0" + upverStr;
                    }
                }
                if (isChange)
                {
                    SucPOIssueNo = mainNo + "CH-" + upverStr;
                }
                if (isCancel)
                {
                    if (poIssueNo == null)
                    {
                        SucPOIssueNo = "-";
                    }
                    else
                    {
                        SucPOIssueNo = mainNo + "CC-" + upverStr;
                    }
                }
                if (isReIsuse)
                {
                    SucPOIssueNo = mainNo + "RE-" + upverStr;
                }

            }
            return SucPOIssueNo;
        }

        public JsonResult UploadFileChangeAndCancel(string action, string reason, HttpPostedFileBase importFile, HttpPostedFileBase evident)
        {

            if (Session["Dept"].ToString() != "PUR")
            {
                return Json("false", JsonRequestBehavior.AllowGet);
            }

            List<string> fail = new List<string>();
            List<string> success = new List<string>();
            var today = DateTime.Now;
            var fileNameEvidential = "";
            if (evident != null) // Luu file Evident
            {
                fileNameEvidential = "Evidential-" + today.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(evident.FileName);
                string path = Server.MapPath("~/File/PO/Evidential/");
                fileNameEvidential += fileExt;
                evident.SaveAs(path + Path.GetFileName(fileNameEvidential));
            }

            if (importFile != null) // Luu file Import
            {
                string fileNameImport = "FileImport-" + action + today.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(importFile.FileName);
                string path = Server.MapPath("~/File/PO/");
                fileNameImport += fileExt;
                importFile.SaveAs(path + Path.GetFileName(fileNameImport));

                // Doc file
                using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/PO/" + fileNameImport))))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    var rowPartNo = 0;
                    var rowDim = 0;
                    var rowECN = 0;
                    var rowHis = 0;
                    var rowDlvKeyNo = 0;

                    var rowDlvDate = 0;
                    // Tìm cột nào là Parts No + Dim + ECN_No
                    for (int i = 1; i <= end.Column; i++)
                    {
                        if (worksheet.Cells[1, i].Text == "Parts No")
                        {
                            rowPartNo = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Dim")
                        {
                            rowDim = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Drawing No.")
                        {
                            rowHis = i;
                        }
                        if (worksheet.Cells[1, i].Text == "ECN No")
                        {
                            rowECN = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Dlv Key No")
                        {
                            rowDlvKeyNo = i;
                        }
                        if (worksheet.Cells[1, i].Text == "Delivery Date")
                        {
                            rowDlvDate = i;
                        }
                    }


                    for (int row = start.Row + 1; row <= end.Row; row++)
                    { // Row by row...
                        var partNo = worksheet.Cells[row, rowPartNo].Text.ToUpper().Trim();
                        var dim = worksheet.Cells[row, rowDim].Text.ToUpper().Trim();
                        var his = worksheet.Cells[row, rowHis].Text.ToUpper().Trim();
                        if (his.Length == 1)
                        {
                            his = "00" + his;
                        }
                        if (his.Length == 2)
                        {
                            his = "0" + his;
                        }
                        var ECN = worksheet.Cells[row, rowECN].Text.ToUpper().Trim();
                        var dlvDate = worksheet.Cells[row, rowDlvDate].Text.ToUpper().Trim();
                        var dlvKeyNo = worksheet.Cells[row, rowDlvKeyNo].Text.ToUpper().Trim();
                        //if (dlvKeyNo.Length == 6)
                        //{
                        //    dlvKeyNo =  dlvKeyNo + "0000";
                        //}
                        //if (dlvKeyNo.Length == 7)
                        //{
                        //    dlvKeyNo = dlvKeyNo + "000";
                        //}
                        //if (dlvKeyNo.Length == 8)
                        //{
                        //    dlvKeyNo = dlvKeyNo + "00";
                        //}
                        //if (dlvKeyNo.Length == 9)
                        //{
                        //    dlvKeyNo = dlvKeyNo + "0";
                        //}
                        if (String.IsNullOrEmpty(partNo)) break;


                        if (action == "uploadFileChange") // change PO
                        {
                            var ExistPO = db.PO_Dies.Where(x => x.MR.PartNo == partNo && x.MR.Clasification == dim
                            && x.MR.DrawHis == his && x.MR.ECNNo == ECN && (x.POStatusID == 8 || x.POStatusID == 15) && x.Active != false).FirstOrDefault();

                            if (ExistPO != null)
                            {
                                if (ExistPO.POstatusCalogory.isCanChangePO == true) // Cho phep change PO khi W-PAyemt (8) || Bi Back lai (15)
                                {
                                    ExistPO.POStatusID = 11; // Change_W-M1-PUR- App
                                    ExistPO.Change_DeliveryDate = Convert.ToDateTime(dlvDate);
                                    ExistPO.Change_DeliveryKey = dlvKeyNo;
                                    ExistPO.Change_Reason = reason;
                                    ExistPO.Change_Evidential = fileNameEvidential;

                                    // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                                    var oldPOIssueNo = ExistPO.POIssueNo;
                                    // Tạo số PO mới
                                    ExistPO.POIssueNo = createPOIssueNo(ExistPO.POIssueNo, false, true, false, false);
                                    ExistPO.Change_RequestBy = Session["Name"].ToString();
                                    ExistPO.Change_RequestDate = today;
                                    // Luu DB
                                    db.Entry(ExistPO).State = EntityState.Modified;
                                    db.SaveChanges();
                                    success.Add(ExistPO.MR.MRNo);
                                }
                                else
                                {
                                    fail.Add(partNo + "-" + dim + ": Ko thể Change PO ở status '" + ExistPO.POstatusCalogory.Status + "'");
                                }
                            }
                            else
                            {
                                fail.Add(partNo + "-" + dim + ": ko tìm thấy PO Die này or Ko thể Change PO ở status nay!");
                            }

                        }

                        if (action == "uploadFileCancel") // cancel PO
                        {
                            var ExistPO = db.PO_Dies.Where(x => x.MR.PartNo == partNo && x.MR.Clasification == dim
                           && x.MR.DrawHis == his && x.MR.ECNNo == ECN && x.Active != false).FirstOrDefault();

                            if (ExistPO != null)
                            {
                                // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                                var oldPOIssueNo = ExistPO.POIssueNo;
                                // Tạo số PO mới
                                ExistPO.POIssueNo = createPOIssueNo(ExistPO.POIssueNo, false, false, true, false);

                                if (ExistPO.POstatusCalogory.isCanCancelPOWithoutRoute == false) // Fải chạy route qua PUC
                                {
                                    ExistPO.POStatusID = 16; // Cancel_W-PUR-M1-APPROVE
                                    ExistPO.Cancel_Reason = reason;
                                    ExistPO.Cancel_DeliveryKey = dlvKeyNo;
                                    ExistPO.Reissue = false;
                                    ExistPO.Cancel_RequestBy = Session["Name"].ToString();
                                    ExistPO.Cancel_RequestDate = today;
                                    ExistPO.Cancel_Evidential = fileNameEvidential;
                                    db.Entry(ExistPO).State = EntityState.Modified;
                                    db.SaveChanges();
                                    success.Add(ExistPO.MR.MRNo);

                                }
                                else
                                {
                                    if (ExistPO.POstatusCalogory.isCanCancelPOWithoutRoute == true)
                                    {
                                        //1. Cancel PO => Ko cần App
                                        ExistPO.POStatusID = 20; // cancel
                                        ExistPO.Cancel_Reason = reason;
                                        ExistPO.Cancel_DeliveryKey = dlvKeyNo;
                                        ExistPO.Cancel_Evidential = fileNameEvidential;
                                        db.Entry(ExistPO).State = EntityState.Modified;
                                        //2. Cancel MR =>
                                        var mR = db.MRs.Find(ExistPO.MRID);
                                        mR.StatusID = 12; // cancel
                                        mR.Note = DateTime.Now.ToString("yyyy-MM-dd: ") + Session["Name"] + " Cancel MR & PO Reason: " + reason + System.Environment.NewLine + mR.Note;
                                        db.Entry(mR).State = EntityState.Modified;
                                        db.SaveChanges();
                                        success.Add(ExistPO.MR.MRNo);
                                    }
                                    else
                                    {
                                        fail.Add(partNo + "-" + dim + ": ko thể cancel PO đang status '" + ExistPO.POstatusCalogory.Status + "''");
                                    }
                                }

                            }
                            else
                            {
                                fail.Add(partNo + "-" + dim + ": ko tìm thấy PO Die này");
                            }

                        }
                        if (action == "uploadFileCancelAndReissue") //Cancel and Reissue
                        {
                            var ExistPO = db.PO_Dies.Where(x => x.MR.PartNo == partNo && x.MR.Clasification == dim
                            && x.MR.DrawHis == his && x.MR.ECNNo == ECN && x.Active != false).FirstOrDefault();

                            if (ExistPO != null)
                            {
                                // Lấy số POissue cũ để trả về view => Mục đích để non-display Item này
                                var oldPOIssueNo = ExistPO.POIssueNo;
                                // Tạo số PO mới
                                ExistPO.POIssueNo = createPOIssueNo(ExistPO.POIssueNo, false, false, true, false);

                                if (ExistPO.POstatusCalogory.isCanCancelPOWithoutRoute == false) // Fải chạy route qua PUC
                                {

                                    ExistPO.POStatusID = 16; // Cancel_W-PUR-M1-APPROVE
                                    ExistPO.Cancel_Reason = reason;
                                    ExistPO.Cancel_DeliveryKey = dlvKeyNo;
                                    ExistPO.Reissue = true;
                                    ExistPO.Cancel_RequestBy = Session["Name"].ToString();
                                    ExistPO.Cancel_RequestDate = today;
                                    ExistPO.Cancel_Evidential = fileNameEvidential;
                                    db.Entry(ExistPO).State = EntityState.Modified;
                                    db.SaveChanges();
                                    success.Add(ExistPO.MR.MRNo);

                                }
                                else
                                {
                                    if (ExistPO.POstatusCalogory.isCanCancelPOWithoutRoute == true)
                                    {
                                        //1. Cancel PO => Ko cần App
                                        ExistPO.POStatusID = 20; // cancel
                                        ExistPO.Cancel_Reason = reason;
                                        ExistPO.Cancel_DeliveryKey = dlvKeyNo;
                                        ExistPO.Cancel_Evidential = fileNameEvidential;
                                        db.Entry(ExistPO).State = EntityState.Modified;


                                        //2. Reissue => Copy New PO
                                        PO_Dies newPOReissue = new PO_Dies();
                                        if (ExistPO.MR.EstimateCost == 10)
                                        {
                                            newPOReissue.TempPO = true;
                                        }
                                        else
                                        {
                                            newPOReissue.TempPO = false;
                                        }
                                        var PR = ExistPO.MR.Clasification.ToString();
                                        newPOReissue.POIssueNo = ExistPO.POIssueNo;
                                        newPOReissue.PR = PR.Remove(PR.Length - 1);
                                        newPOReissue.MRID = ExistPO.MR.MRID;
                                        newPOReissue.POStatusID = 21; // W-PUR-Re-ISSUE
                                        newPOReissue.Active = true;
                                        newPOReissue.DeliveryDate = ExistPO.MR.PDD;
                                        newPOReissue.CreateDate = today;
                                        db.PO_Dies.Add(newPOReissue);
                                        db.SaveChanges();
                                        success.Add(ExistPO.MR.MRNo);
                                    }
                                    else
                                    {
                                        fail.Add(partNo + "-" + dim + ": ko thể cancel PO đang status '" + ExistPO.POstatusCalogory.Status + "''");
                                    }
                                }

                            }
                            else
                            {
                                fail.Add(partNo + "-" + dim + ": ko tìm thấy PO Die này");
                            }

                        }
                    }

                }

            }



            var data = new
            {
                success,
                fail
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        // Temp


        //public void updateDieLaunching(PO_Dies po)
        // {
        //     try
        //     {
        //         var partNo = po.MR.PartNo;
        //         var dieNo = po.MR.Clasification;
        //         var result = commonFunction.isRenewOrAddOrMT(dieNo);
        //         if (result.Contains("Renew") || result.Contains("Add"))
        //         {
        //             var existDieLaunching = db.Die_Launch_Management.Where(x => x.Part_No == partNo && x.Die_No == dieNo && x.isActive != false).FirstOrDefault();
        //             if (existDieLaunching != null) // Null thi add new
        //             {
        //                 existDieLaunching.PO_Issue_Date = existDieLaunching.PO_Issue_Date == null ? po.IssueDate : existDieLaunching.PO_Issue_Date;
        //                 existDieLaunching.PO_App_Date = existDieLaunching.PO_App_Date == null ? po.PODate : existDieLaunching.PO_App_Date;
        //                 db.Entry(existDieLaunching).State = EntityState.Modified;
        //                 db.SaveChanges();
        //             }
        //         }
        //     }
        //     catch
        //     {

        //     }
        // }
        public string handelDeliveryKeyNo(string key)
        {
            string output = "";
            if (!String.IsNullOrWhiteSpace(key))
            {

                // Nếu key ko tồn tại dấu cách và length = 10.
                if (key.Length == 10 && !key.Contains(" "))
                {
                    output = key;
                }
                else
                {
                    //1. Loại bỏ toàn bộ dấu cách "123456_0" && "123456___0
                    try
                    {
                        output = key.Substring(0, 6);
                    }
                    catch
                    {
                        output = "Fail";
                    }
                }

            }


            return output;
        }

        //public int createDie(int mRTypeID, int mRID)
        //{
        //    var dieIDSuccess = 0;
        //    gen
        //    switch (mRTypeID)
        //    {
        //        case 1:
        //            var dieID = createNewDie(mRID);
        //            if (dieID > 0)
        //            {
        //                dieIDSuccess = dieID;
        //                sendEmailToAdmin(dieID, mRID, "OK");

        //            }
        //            else
        //            {
        //                sendEmailToAdmin(dieID, mRID, "NG");

        //            }
        //            break;
        //        case 2:
        //            var dieID2 = CreateAdditional(mRID);
        //            if (dieID2 > 0)
        //            {
        //                dieIDSuccess = dieID2;
        //                sendEmailToAdmin(dieID2, mRID, "OK");

        //            }
        //            else
        //            {
        //                sendEmailToAdmin(dieID2, mRID, "NG");

        //            }
        //            break;
        //        case 3:
        //            var dieID3 = CreateRenewal(mRID);
        //            if (dieID3 > 0)
        //            {
        //                dieIDSuccess = dieID3;
        //                sendEmailToAdmin(dieID3, mRID, "OK");
        //            }
        //            else
        //            {
        //                sendEmailToAdmin(dieID3, mRID, "NG");
        //            }
        //            break;

        //    }
        //    return (dieIDSuccess);
        //}
        //public int createNewDie(int mRID)
        //{
        //    var mR = db.MRs.AsNoTracking().Where(x => x.MRID == mRID).FirstOrDefault();
        //    //New Die
        //    Die1 newDie = new Die1();
        //    newDie.DieNo = mR.DieNo;
        //    newDie.PartNoOriginal = mR.PartNo;
        //    newDie.ProcessCodeID = mR.ProcessCodeID.Value;
        //    newDie.ModelID = mR.ModelID.Value;
        //    newDie.DieMaker = db.Suppliers.Where(x => x.SupplierCode == mR.OrderTo).FirstOrDefault().SupplierName;
        //    newDie.SupplierID = mR.SupplierID.Value;
        //    newDie.MCsize = mR.MCSize;
        //    newDie.CavQuantity = mR.CavQty;
        //    newDie.DieCost_USD = mR.AppCostExchangeUSD;
        //    newDie.SpecialSpec = mR.DieSpecial;
        //    newDie.PODate = mR.PODate;
        //    newDie.DieClassify = "MT";
        //    newDie.Active = true;
        //    newDie.Belong = mR.Belong;

        //    try
        //    {
        //        db.Die1.Add(newDie);
        //        db.SaveChanges();
        //        DieStatusUpdateRegular newUpdate = new DieStatusUpdateRegular();
        //        newUpdate.DieID = newDie.DieID;
        //        newUpdate.DieNo = newDie.DieNo;
        //        newUpdate.DieStatus = "Making";
        //        newUpdate.DieOperationStatus = mR.PODate + "  was Approved PO";
        //        newUpdate.IssueDate = DateTime.Now;
        //        newUpdate.IssueBy = Session["Name"].ToString();
        //        db.DieStatusUpdateRegulars.Add(newUpdate);
        //        db.SaveChanges();
        //    }
        //    catch
        //    {
        //        return 0;
        //    }
        //    if (mR.NoOfDieComponent > 1)
        //    {
        //        var dieNo = mR.DieNo; // RC5-1234-000-11A-001
        //        var mainDieNo = dieNo.Remove(dieNo.Length - 3); // RC5-1234-000-11A-
        //        for (var i = 2; i <= mR.NoOfDieComponent; i++)
        //        {
        //            var tail = i;
        //            var tailStr = "";
        //            if (tail.ToString().Length == 1)
        //            {
        //                tailStr = "00" + tail.ToString();
        //            }
        //            if (tail.ToString().Length == 2)
        //            {
        //                tailStr = "0" + tail.ToString();
        //            }

        //            Die1 newDie1 = new Die1();
        //            newDie1 = newDie;
        //            newDie1.DieNo = mainDieNo + tailStr;
        //            db.Die1.Add(newDie1);
        //            db.SaveChanges();
        //        }
        //    }
        //    //new part
        //    // Check part da ton tai trong Database chua
        //    var existPartNo = db.Parts1.Where(x => x.PartNo.Contains(mR.PartNo.Trim())).FirstOrDefault();
        //    int newPartID = 0;
        //    if (existPartNo == null)
        //    {
        //        Parts1 newPart = new Parts1();
        //        newPart.PartNo = mR.PartNo;
        //        newPart.PartName = mR.PartName;
        //        newPart.Model = mR.ModelList.ModelName;
        //        newPart.Active = true;
        //        try
        //        {
        //            db.Parts1.Add(newPart);
        //            db.SaveChanges();
        //            newPartID = newPart.PartID;
        //        }
        //        catch
        //        {
        //            db.Die1.Remove(newDie);
        //            db.SaveChanges();
        //            return 0;
        //        }
        //    }

        //    //New Common
        //    CommonDie1 newCommon = new CommonDie1();
        //    newCommon.DieNo = newDie.DieNo;
        //    newCommon.PartNo = mR.PartNo.Trim().ToUpper();
        //    newCommon.DieID = newDie.DieID;
        //    newCommon.PartID = newPartID;
        //    newCommon.Active = true;
        //    try
        //    {
        //        db.CommonDie1.Add(newCommon);
        //        db.SaveChanges();

        //    }
        //    catch
        //    {
        //        db.Die1.Remove(newDie);
        //        db.SaveChanges();
        //        return 0;
        //    }

        //    //New Part has Family common
        //    if (mR.CommonPart != null)
        //    {
        //        if (mR.CommonPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.CommonPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Common with" + mR.PartNo;
        //                newPartCommonItem.Model = mR.ModelList.ModelName;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }
        //    }
        //    if (mR.FamilyPart != null)
        //    {
        //        if (mR.FamilyPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.FamilyPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Family with" + mR.PartNo;
        //                newPartCommonItem.Model = mR.ModelList.ModelName;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }
        //    }

        //    return (newDie.DieID);
        //}
        //public int CreateAdditional(int mRID)
        //{
        //    var mR = db.MRs.AsNoTracking().Where(x => x.MRID == mRID).FirstOrDefault();
        //    //New Die 
        //    Die1 newDie = new Die1();
        //    newDie.DieNo = mR.DieNo;
        //    newDie.PartNoOriginal = mR.PartNo;
        //    newDie.ProcessCodeID = mR.ProcessCodeID.Value;
        //    newDie.ModelID = mR.ModelID.Value;
        //    newDie.DieMaker = db.Suppliers.Where(x => x.SupplierCode == mR.OrderTo).FirstOrDefault().SupplierName;
        //    newDie.SupplierID = mR.SupplierID.Value;
        //    newDie.MCsize = mR.MCSize;
        //    newDie.CavQuantity = mR.CavQty;
        //    newDie.DieCost_USD = mR.AppCostExchangeUSD;
        //    newDie.SpecialSpec = mR.DieSpecial;
        //    newDie.PODate = mR.PODate;
        //    newDie.DieClassify = "RD&AD";
        //    newDie.Active = true;
        //    newDie.Belong = mR.Belong;

        //    try
        //    {
        //        db.Die1.Add(newDie);
        //        db.SaveChanges();
        //        DieStatusUpdateRegular newUpdate = new DieStatusUpdateRegular();
        //        newUpdate.DieID = newDie.DieID;
        //        newUpdate.DieNo = newDie.DieNo;
        //        newUpdate.DieStatus = "Making";
        //        newUpdate.DieOperationStatus = mR.PODate + "  was Approved PO";
        //        newUpdate.IssueDate = DateTime.Now;
        //        newUpdate.IssueBy = Session["Name"].ToString();
        //        db.DieStatusUpdateRegulars.Add(newUpdate);
        //        db.SaveChanges();
        //    }
        //    catch
        //    {
        //        return 0;
        //    }
        //    if (mR.NoOfDieComponent > 1)
        //    {
        //        var dieNo = mR.DieNo; // RC5-1234-000-11A-001
        //        var mainDieNo = dieNo.Remove(dieNo.Length - 3); // RC5-1234-000-11A-
        //        for (var i = 2; i >= mR.NoOfDieComponent; i++)
        //        {
        //            var tail = i;
        //            var tailStr = "";
        //            if (tail.ToString().Length == 1)
        //            {
        //                tailStr = "00" + tail.ToString();
        //            }
        //            if (tail.ToString().Length == 2)
        //            {
        //                tailStr = "0" + tail.ToString();
        //            }

        //            Die1 newDie1 = new Die1();
        //            newDie1 = newDie;
        //            newDie1.DieNo = mainDieNo + tailStr;
        //            db.Die1.Add(newDie1);
        //            db.SaveChanges();
        //        }
        //    }
        //    //old part
        //    var oldPart = db.Parts1.Where(x => x.PartNo == mR.PartNo.ToUpper().Trim()).FirstOrDefault();
        //    //New Common
        //    CommonDie1 newCommon = new CommonDie1();
        //    newCommon.DieNo = newDie.DieNo;
        //    newCommon.PartNo = oldPart.PartNo;
        //    newCommon.DieID = newDie.DieID;
        //    newCommon.PartID = oldPart.PartID;
        //    newCommon.Active = true;
        //    try
        //    {
        //        db.CommonDie1.Add(newCommon);
        //        db.SaveChanges();
        //    }
        //    catch
        //    {
        //        db.Die1.Remove(newDie);
        //        db.SaveChanges();
        //        return 0;
        //    }
        //    //New Part has Family common
        //    if (mR.CommonPart != null)
        //    {
        //        if (mR.CommonPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.CommonPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Common with" + oldPart.PartNo;
        //                newPartCommonItem.Model = oldPart.Model;
        //                newPartCommonItem.Material = oldPart.Material;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }

        //    }
        //    if (mR.FamilyPart != null)
        //    {
        //        if (mR.FamilyPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.FamilyPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Family with" + oldPart.PartNo;
        //                newPartCommonItem.Model = oldPart.Model;
        //                newPartCommonItem.Material = oldPart.Material;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }

        //    }

        //    return (newDie.DieID);
        //}
        //public int CreateRenewal(int mRID)
        //{
        //    var mR = db.MRs.AsNoTracking().Where(x => x.MRID == mRID).FirstOrDefault();
        //    //New Die 
        //    Die1 newDie = new Die1();
        //    newDie.DieNo = mR.DieNo;
        //    newDie.PartNoOriginal = mR.PartNo;
        //    newDie.ProcessCodeID = mR.ProcessCodeID.Value;
        //    newDie.ModelID = mR.ModelID.Value;
        //    newDie.DieMaker = db.Suppliers.Where(x => x.SupplierCode == mR.OrderTo).FirstOrDefault().SupplierName;
        //    newDie.SupplierID = mR.SupplierID.Value;
        //    newDie.MCsize = mR.MCSize;
        //    newDie.CavQuantity = mR.CavQty;
        //    newDie.DieCost_USD = mR.AppCost;
        //    newDie.SpecialSpec = mR.DieSpecial;
        //    newDie.PODate = mR.PODate;
        //    newDie.DieClassify = "RD&AD";
        //    newDie.Active = true;
        //    newDie.Belong = mR.Belong;

        //    try
        //    {
        //        db.Die1.Add(newDie);
        //        db.SaveChanges();
        //        DieStatusUpdateRegular newUpdate = new DieStatusUpdateRegular();
        //        newUpdate.DieID = newDie.DieID;
        //        newUpdate.DieNo = newDie.DieNo;
        //        newUpdate.DieStatus = "Making";
        //        newUpdate.DieOperationStatus = mR.PODate + "  was Approved PO";
        //        newUpdate.IssueDate = DateTime.Now;
        //        newUpdate.IssueBy = Session["Name"].ToString();
        //        db.DieStatusUpdateRegulars.Add(newUpdate);
        //        db.SaveChanges();
        //    }
        //    catch
        //    {
        //        return 0;
        //    }

        //    if (mR.NoOfDieComponent > 1)
        //    {
        //        var dieNo = mR.DieNo; // RC5-1234-000-11A-001
        //        var mainDieNo = dieNo.Remove(dieNo.Length - 3); // RC5-1234-000-11A-
        //        for (var i = 2; i >= mR.NoOfDieComponent; i++)
        //        {
        //            var tail = i;
        //            var tailStr = "";
        //            if (tail.ToString().Length == 1)
        //            {
        //                tailStr = "00" + tail.ToString();
        //            }
        //            if (tail.ToString().Length == 2)
        //            {
        //                tailStr = "0" + tail.ToString();
        //            }

        //            Die1 newDie1 = new Die1();
        //            newDie1 = newDie;
        //            newDie1.DieNo = mainDieNo + tailStr;
        //            db.Die1.Add(newDie1);
        //            db.SaveChanges();
        //        }
        //    }
        //    //old part
        //    var oldPart = db.Parts1.Where(x => x.PartNo == mR.PartNo.ToUpper().Trim()).FirstOrDefault();
        //    //New Common
        //    CommonDie1 newCommon = new CommonDie1();
        //    newCommon.DieNo = newDie.DieNo;
        //    newCommon.PartNo = oldPart.PartNo;
        //    newCommon.DieID = newDie.DieID;
        //    newCommon.PartID = oldPart.PartID;
        //    newCommon.Active = true;
        //    try
        //    {
        //        db.CommonDie1.Add(newCommon);
        //        db.SaveChanges();
        //    }
        //    catch
        //    {
        //        db.Die1.Remove(newDie);
        //        db.SaveChanges();
        //        return 0;
        //    }
        //    //New Part has Family common
        //    if (mR.CommonPart != null)
        //    {
        //        if (mR.CommonPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.CommonPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Common with" + oldPart.PartNo;
        //                newPartCommonItem.Model = oldPart.Model;
        //                newPartCommonItem.Material = oldPart.Material;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }

        //    }
        //    if (mR.FamilyPart != null)
        //    {
        //        if (mR.FamilyPart.Trim().Length > 7)
        //        {
        //            string[] arrListPart = mR.FamilyPart.Split(',');
        //            foreach (var item in arrListPart)
        //            {
        //                Parts1 newPartCommonItem = new Parts1();
        //                newPartCommonItem.PartNo = item.Trim().ToUpper();
        //                newPartCommonItem.Note = "Family with" + oldPart.PartNo;
        //                newPartCommonItem.Model = oldPart.Model;
        //                newPartCommonItem.Material = oldPart.Material;
        //                newPartCommonItem.Active = true;
        //                db.Parts1.Add(newPartCommonItem);
        //                db.SaveChanges();

        //                CommonDie1 newCommonItem = new CommonDie1();
        //                newCommonItem.DieNo = newDie.DieNo;
        //                newCommonItem.DieID = newDie.DieID;
        //                newCommonItem.PartNo = newPartCommonItem.PartNo;
        //                newCommonItem.PartID = newPartCommonItem.PartID;
        //                newCommonItem.Active = true;
        //                db.CommonDie1.Add(newCommonItem);
        //                db.SaveChanges();
        //            }
        //        }

        //    }
        //    //Update old die to "Has Renew"
        //    var oldDie = db.Die1.Where(x => x.DieNo == mR.RenewForDie.ToUpper().Trim()).FirstOrDefault();
        //    oldDie.DieClassify = "Had Renew";
        //    db.Entry(oldDie).State = EntityState.Modified;
        //    db.SaveChanges();

        //    return (newDie.DieID);
        //}
        public int handelSuccesDie(int mRID)
        {
            int value = 0;
            try
            {
                var mR = db.MRs.AsNoTracking().Where(x => x.MRID == mRID).FirstOrDefault();
                var succesDieNo = mR.SucessDieID;
                var succesPartName = mR.SucessPartName;
                var currentDie = db.Die1.Where(x => x.DieNo == mR.DieNo).FirstOrDefault();
                currentDie.SpecialSpec = "This Die was change to " + mR.SucessDieID + "by MRNo: " + mR.MRNo;
                db.Entry(currentDie).State = EntityState.Modified;

                Die1 newDie = new Die1();
                newDie = currentDie;
                newDie.DieNo = mR.SucessDieID.Trim().ToUpper();
                newDie.PartNoOriginal = mR.SucessPartNo;
                newDie.SpecialSpec = "This Die was changed from " + mR.DieNo + " to " + mR.SucessDieID + " by MRNo: " + mR.MRNo;
                db.Die1.Add(newDie);
                db.SaveChanges();

                var successPartNo = db.Parts1.Where(x => x.PartNo == mR.SucessPartNo).FirstOrDefault();
                if (successPartNo == null)
                {
                    Parts1 newPart = new Parts1();
                    newPart.PartNo = mR.SucessPartNo.Trim();
                    newPart.Note = "Common/Family with" + mR.PartNo;
                    newPart.Model = mR.ModelName;
                    newPart.Material = db.Parts1.Where(x => x.PartNo == mR.PartNo.Trim()).FirstOrDefault().Material;
                    newPart.Active = true;
                    db.Parts1.Add(newPart);
                    db.SaveChanges();
                    successPartNo = newPart;
                }

                CommonDie1 newCommon = new CommonDie1();
                newCommon.DieNo = newDie.DieNo;
                newCommon.PartNo = successPartNo.PartNo;
                newCommon.DieID = newDie.DieID;
                newCommon.PartID = successPartNo.PartID;
                newCommon.Active = true;
                db.CommonDie1.Add(newCommon);
                db.SaveChanges();

                if (!String.IsNullOrEmpty(mR.CommonPart))
                {
                    handelCommonOrFaminlyPart(newDie.DieNo, mR.CommonPart, mR.ModelName);
                }
                value = 1;
            }
            catch
            {
                value = 0;
            }

            return value;
        }
        public bool cancelDie_Part_Common(int? MRID)
        {
            bool status = false;
            var mR = db.MRs.Find(MRID);
            if (mR.TypeID == 1 || mR.TypeID == 2 || mR.TypeID == 3)
            {
                var Commom = db.CommonDie1.Where(x => x.DieNo == mR.DieNo).ToList();
                var dies = db.Die1.Where(x => x.DieNo == mR.DieNo).ToList();

                foreach (var item in dies)
                {
                    item.Active = true;
                    item.isOfficial = false;
                    item.Genaral_Information = DateTime.Now.ToString("MM/dd/yyyy") + " :PUR cancel PO" + System.Environment.NewLine + item.Genaral_Information;
                    db.Entry(item).State = EntityState.Modified;
                }
                db.SaveChanges();
                status = true;
            }
            return status;
        }
        public int handelCommonOrFaminlyPart(string dieNo, string comompartStringList, string modelName)
        {
            var value = 0;
            try
            {
                var checkDie = db.Die1.Where(x => x.DieNo == dieNo.Trim().ToUpper()).FirstOrDefault();
                // comompartStringList = "RC5-1234-000,RC5-4567-000"
                string[] arrListPart = comompartStringList.Split(',');
                foreach (var item in arrListPart)
                {
                    var partNo = item.Trim().ToUpper();
                    var checkPartExit = db.Parts1.Where(x => x.PartNo.Contains(partNo)).FirstOrDefault();
                    if (checkPartExit == null)
                    {
                        Parts1 newPartCommonItem = new Parts1();
                        newPartCommonItem.PartNo = partNo;
                        newPartCommonItem.Note = "Common/Family with" + comompartStringList;
                        newPartCommonItem.Model = modelName;
                        newPartCommonItem.Active = true;
                        db.Parts1.Add(newPartCommonItem);
                        db.SaveChanges();
                        checkPartExit = newPartCommonItem;
                    }

                    CommonDie1 newCommonItem = new CommonDie1();
                    newCommonItem.DieNo = checkDie.DieNo;
                    newCommonItem.DieID = checkDie.DieID;
                    newCommonItem.PartNo = checkPartExit.PartNo;
                    newCommonItem.PartID = checkPartExit.PartID;
                    newCommonItem.Active = true;
                    db.CommonDie1.Add(newCommonItem);
                    db.SaveChanges();
                    value = 1;
                }
            }
            catch
            {
                value = 0;
            }
            return value;
        }
        public void sendEmailToAdmin(int dieID, int mRID, string content)
        {
            var subject = "";
            if (content == "OK")
            {
                var die = db.Die1.Find(dieID);
                subject = "System has just create new die No: " + die.DieNo;
            }
            else
            {
                var mR = db.MRs.Find(mRID);
                subject = "System fail to create new die from MR No: " + mR.MRNo;
            }
            var Admin = db.Users.Where(x => x.Department.DeptName == "PAE" && x.Role == "Admin" && x.Active == true);
            var mailList = "";
            foreach (var item in Admin)
            {
                string mail = item.Email;
                mailList = mailList + "," + mail;
            }
            try
            {
                MailMessage mailMsg = new MailMessage();
                mailMsg.From = new MailAddress("QV-DMS@canon-vn.com.vn");
                mailMsg.To.Add(mailList);

                mailMsg.Subject = subject;
                mailMsg.IsBodyHtml = true;
                mailMsg.Body = (
                    " Dear PAE Admin of DMS " + " <br /> " +
                                       "<br />" +
                                             "<br />" +
                                                "Please check content:" + "<br />" +
                                                subject +

                                             "<br />" +
                                          "P/s: Please don't reply this email <br />" +
                                          "***************************** <br />" +
                                          "Thanks & Best Regards!"
                    );
                mailMsg.Priority = MailPriority.Normal;
                SmtpClient client = new SmtpClient("mail.cvn.canon.co.jp", 2525)
                {
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential("lbp-pae21@canon-vn.com.vn", ""),
                    EnableSsl = false
                };
                client.UseDefaultCredentials = false;
                ThreadStart threadStart = delegate () { client.Send(mailMsg); };
                Thread thread = new Thread(threadStart);
                thread.Start();
            }
            catch
            {
                //
            }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
