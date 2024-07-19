using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Authentication.ExtendedProtection;
using System.Web;
using System.Web.Mvc;
using DMS03.Models;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PagedList;

namespace DMS03.Controllers
{
    public class DieTransferController : Controller
    {
        private DMSEntities db = new DMSEntities();
        private SendEmailController sendEmailJob = new SendEmailController();
        // GET: DieTransfer
        public ActionResult Index(int? page,string search, string from, string to, string export, string waitfor, string showAll)
        {
            if (page == null) page = 1;
            int pageSize = 50;
            int pageNumber = (page ?? 1);

            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            var allRQ = db.DieLendingRequests.Where(x => x.Active != false).ToList();
            var searchResult = allRQ;


            if (!String.IsNullOrEmpty(search))
            {
                searchResult = searchResult.Where(x => x.DieNo.Contains(search)).ToList();
                if (searchResult.Count() == 0)
                {
                    searchResult = allRQ.Where(x => x.FixedAssetNo != null && x.FixedAssetNo.Contains(search)).ToList();
                }
                if (searchResult.Count() == 0)
                {
                    searchResult = allRQ.Where(x => x.DTFNo.Contains(search)).ToList();
                }
            }
            if (!String.IsNullOrEmpty(from))
            {
                searchResult = searchResult.Where(x => x.RequestDate >= DateTime.Parse(from)).ToList();
            }
            if (!String.IsNullOrEmpty(to))
            {
                searchResult = searchResult.Where(x => x.RequestDate <= DateTime.Parse(to)).ToList();
            }
            if (!String.IsNullOrEmpty(waitfor))
            {
                searchResult = searchResult.Where(x => x.LendingStatusCategory.Type.ToLower().Contains(waitfor.ToLower())).ToList();
            }
            if (!String.IsNullOrEmpty(showAll))
            {
                 searchResult = allRQ;
            }

            if (export != null)
            {
                exportToList(searchResult);
            }

            ViewBag.search = search;
            ViewBag.from = from;
            ViewBag.to = to;
            ViewBag.W_DMT_Check = allRQ.Where(x => x.LendingStatusID == 1).Count();
            ViewBag.W_DMT_Approve = allRQ.Where(x => x.LendingStatusID == 2).Count();
            ViewBag.W_PUR_Check = allRQ.Where(x => x.LendingStatusID == 3).Count();
            ViewBag.W_PUR_Approve = allRQ.Where(x => x.LendingStatusID == 4).Count();
            ViewBag.W_PUC_Check = allRQ.Where(x => x.LendingStatusID == 5).Count();
            ViewBag.W_PUC_Approve = allRQ.Where(x => x.LendingStatusID == 6).Count();
            ViewBag.W_PAE_Check = allRQ.Where(x => x.LendingStatusID == 7).Count();
            ViewBag.W_PAE_Approve = allRQ.Where(x => x.LendingStatusID == 8).Count();
            ViewBag.Rejected = allRQ.Where(x => x.LendingStatusID == 11).Count();
            ViewBag.Cancelled = allRQ.Where(x => x.LendingStatusID == 12).Count();

            ViewBag.SupplierID = new SelectList(db.Suppliers, "SupplierID", "SupplierCode");

            return View(searchResult.ToPagedList(pageNumber, pageSize));
        }


        public ActionResult IssueLendingRequest(int? id)
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            //****************************
            //****************************
            //****************************
            var dept = Session["Dept"].ToString();
            var role = Session["Lending_Role"].ToString();
            var err = "";

            if (id != null)
            {
                var rq = db.DieLendingRequests.Find(id);
                if (rq.LendingStatusCategory.IsCanIssue == true && (role == "Check" || role == "Approve") && dept == rq.RequestDept)
                {
                    ViewBag.id = id;
                }
                else
                {
                    if (rq.LendingStatusCategory.IsCanIssue != true)
                    {
                        err = "This request is processing, You can not revise it!";
                    }
                    else
                    {
                        if (role != "Check" && role != "Approve")
                        {
                            err = "You do not have permision issue/Revise Trasfer request!";
                        }
                        else
                        {
                            err = "You (" + dept + ") can not issue/Revise request of " + rq.RequestDept;
                        }
                    }
                }

            }

            ViewBag.err = err;
            ViewBag.SupplierID = new SelectList(db.Suppliers, "SupplierID", "SupplierCode");



            return View();
        }

        //public JsonResult getAutoFill(string fixedAssetNo)
        //{
        //    db.Configuration.ProxyCreationEnabled = false;
        //    var existLending = db.DieLendingItems.Where(x => x.FixAssetNo == fixedAssetNo && x.DieLendingRequest.LendingTypeID == 1).OrderByDescending(x => x.LendingItemID).FirstOrDefault(); // "==1" Type= Lending
        //    db.Configuration.ProxyCreationEnabled = false;
        //    var existFixedAsset = db.Die1.Where(x => x.FixedAssetNo == fixedAssetNo && x.Active != false).AsNoTracking().FirstOrDefault();
        //    Object FixedAsset = null;
        //    if (existFixedAsset != null)
        //    {
        //        FixedAsset = new
        //        {
        //            PartNo = existFixedAsset != null ? existFixedAsset.PartNoOriginal : "",
        //            PartName = existFixedAsset != null ? db.Parts1.Where(x => x.PartNo == existFixedAsset.PartNoOriginal).FirstOrDefault().PartName : "",
        //            DieNo = existFixedAsset != null ? existFixedAsset.DieNo : "",
        //            ModelName = existFixedAsset != null ? db.ModelLists.Find(existFixedAsset.ModelID).ModelName : ""
        //        };
        //    }

        //    var data = new
        //    {
        //        Lending = existLending,
        //        FixedAsset = FixedAsset
        //    };
        //    return Json(data, JsonRequestBehavior.AllowGet);
        //}


        public JsonResult issueDTF(string LendingType, string FixAssetNo, string dieNo, string actualShot, string ETAPlan, string ETDPlan, string Transportation, string CurrentLocation, string NewLocation, string Remark)
        {
            
            var today = DateTime.Now;
            var status = false;
            var msg = "";
            // check enough data or not?
            if (String.IsNullOrWhiteSpace(LendingType) || String.IsNullOrWhiteSpace(FixAssetNo) || String.IsNullOrWhiteSpace(dieNo) ||
              String.IsNullOrWhiteSpace(actualShot) || String.IsNullOrWhiteSpace(ETAPlan) || String.IsNullOrWhiteSpace(ETDPlan) ||
                String.IsNullOrWhiteSpace(Transportation) || String.IsNullOrWhiteSpace(CurrentLocation) || String.IsNullOrWhiteSpace(NewLocation))
            {
                status = false;
                msg = "Please input enough information!";
                goto exit;
            }

            // Kiem tra da co tồn tại request đang trong tiến trình xử lí ko?
            // Nếu đã tồn tại => ko cho phép issue
            // Nếu chưa tồn tại => được phép issue
            // ****
            var isExistRequest = db.DieLendingRequests.Where(x => x.FixedAssetNo == FixAssetNo.ToUpper() && x.Active != false && x.LendingStatusID <= 9).FirstOrDefault();
            if (isExistRequest != null)
            {
                status = false;
                msg = "Already exist request trasfer on processing DFTNo: " + isExistRequest.DTFNo;
                goto exit;
            }

            var checkDieExist = db.Die1.Where(x => x.DieNo == dieNo.ToUpper() && x.Active != false).FirstOrDefault();
            if (checkDieExist == null)
            {
                status = false;
                msg = "Die ID is not exist, Pls re-check it";
                goto exit;
            }

            //***
            //1. Tạo new request on DB
            DieLendingRequest newRQ = new DieLendingRequest();
            newRQ.LendingTypeID = int.Parse(LendingType);
            newRQ.FixedAssetNo = FixAssetNo.ToUpper();
            newRQ.DieNo = dieNo.ToUpper();
            newRQ.ActualShot = float.Parse(actualShot);
            newRQ.ETAPlan = DateTime.Parse(ETAPlan);
            newRQ.ETDPlan = DateTime.Parse(ETDPlan);
            newRQ.Transport = Transportation;
            newRQ.CurrentLocation = db.Suppliers.Find(int.Parse(CurrentLocation)).SupplierCode;
            newRQ.NewLocation = db.Suppliers.Find(int.Parse(NewLocation)).SupplierCode;
            newRQ.CurrentLocationID = int.Parse(CurrentLocation);
            newRQ.NewLocationID = int.Parse(NewLocation);
            newRQ.ModelName = checkDieExist.ModelList.ModelName;
            newRQ.Remark = Remark;

            //2. Tạo số DTF
            newRQ.DTFNo = genarateDTFNo("", true, false);
            //3. update progress and next status
            newRQ = genarateProgress(newRQ);

            // Lưu
            db.DieLendingRequests.Add(newRQ);
            db.SaveChanges();
            sendEmailJob.sendEmailDieTransfer(newRQ);
            status = true;
            msg = "OK";
        exit:
            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }


        public ActionResult issueDTFByExcel(HttpPostedFileBase file)
        {

            var today = DateTime.Now;
            string fileName = "ResultIssueDTF-" + today.ToString("yyyy-MM-dd-hhss");
            string fileExt = Path.GetExtension(file.FileName);
            if (fileExt == ".xls" || fileExt == ".xlsx")
            {

                MemoryStream output = new MemoryStream();
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    for (int row = start.Row + 2; row <= end.Row; row++)
                    {
                        var dieNo = worksheet.Cells[row, 4].Text.Trim().ToUpper();
                        if (String.IsNullOrEmpty(dieNo)) break;
                        // (string LendingType, string FixAssetNo, string dieNo, string actualShot, string ETAPlan, string ETDPlan, string Transportation, string CurrentLocation, string NewLocation, string Remark
                        var LendingType = worksheet.Cells[row, 3].Text.Trim();
                        var FixAssetNo = worksheet.Cells[row, 5].Text.Trim();

                        var actualShot = worksheet.Cells[row, 8].Text.Trim();
                        var ETAPlan = worksheet.Cells[row, 10].Text.Trim();
                        var ETDPlan = worksheet.Cells[row, 9].Text.Trim();
                        var Transportation = worksheet.Cells[row, 11].Text.Trim();
                        var CurrentLocationCode = worksheet.Cells[row, 6].Text.Trim().ToUpper();
                        var NewLocationCode = worksheet.Cells[row, 7].Text.Trim().ToUpper();
                        var Remark = worksheet.Cells[row, 12].Text.Trim();
                        var CurrentSupplier = db.Suppliers.Where(x => x.SupplierCode == CurrentLocationCode).FirstOrDefault();
                        var NewSupplier = db.Suppliers.Where(x => x.SupplierCode == NewLocationCode).FirstOrDefault();
                        var CurrentLocationID = CurrentSupplier == null ? "" : CurrentSupplier.SupplierID.ToString();
                        var NewLocationID = NewSupplier == null ? "" : NewSupplier.SupplierID.ToString();

                        var result = issueDTF(LendingType, FixAssetNo, dieNo, actualShot, ETAPlan, ETDPlan, Transportation, CurrentLocationID, NewLocationID, Remark);
                        worksheet.Cells[row, 13].Value = result.Data.ToString();
                    }
                    package.SaveAs(output);
                }
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".xlsx");

                output.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
            return RedirectToAction("Index");
        }

        public JsonResult reviseDTF(string action,string LendingType, string FixAssetNo, string dieNo, string actualShot, string ETAPlan, string ETDPlan, string Transportation, string CurrentLocation, string NewLocation, string Remark)
        {

            var today = DateTime.Now;
            var status = false;
            var msg = "";
            // check enough data or not?
            if (String.IsNullOrWhiteSpace(LendingType) || String.IsNullOrWhiteSpace(FixAssetNo) || String.IsNullOrWhiteSpace(dieNo) ||
              String.IsNullOrWhiteSpace(actualShot) || String.IsNullOrWhiteSpace(ETAPlan) || String.IsNullOrWhiteSpace(ETDPlan) ||
                String.IsNullOrWhiteSpace(Transportation) || String.IsNullOrWhiteSpace(CurrentLocation) || String.IsNullOrWhiteSpace(NewLocation))
            {
                status = false;
                msg = "Please input enough information!";
                goto exit;
            }

            // Kiem tra da co tồn tại request đang trong tiến trình xử lí ko?
            // Nếu đã tồn tại => ko cho phép issue
            // Nếu chưa tồn tại => được phép issue
            // ****
            var isExistRequest = db.DieLendingRequests.Where(x => x.FixedAssetNo == FixAssetNo.ToUpper() && x.Active != false && x.LendingStatusID <= 9).FirstOrDefault();
            if (isExistRequest != null)
            {
                status = false;
                msg = "Already exist request trasfer on processing DFTNo: " + isExistRequest.DTFNo;
                goto exit;
            }

            var checkDieExist = db.Die1.Where(x => x.DieNo == dieNo.ToUpper() && x.Active != false).FirstOrDefault();
            if (checkDieExist == null)
            {
                status = false;
                msg = "Die ID is not exist, Pls re-check it";
                goto exit;
            }

            //***
            //1. Tạo new request on DB
            DieLendingRequest newRQ = new DieLendingRequest();
            newRQ.LendingTypeID = int.Parse(LendingType);
            newRQ.FixedAssetNo = FixAssetNo.ToUpper();
            newRQ.DieNo = dieNo.ToUpper();
            newRQ.ActualShot = float.Parse(actualShot);
            newRQ.ETAPlan = DateTime.Parse(ETAPlan);
            newRQ.ETDPlan = DateTime.Parse(ETDPlan);
            newRQ.Transport = Transportation;
            newRQ.CurrentLocation = db.Suppliers.Find(int.Parse(CurrentLocation)).SupplierCode;
            newRQ.NewLocation = db.Suppliers.Find(int.Parse(NewLocation)).SupplierCode;
            newRQ.CurrentLocationID = int.Parse(CurrentLocation);
            newRQ.NewLocationID = int.Parse(NewLocation);
            newRQ.ModelName = checkDieExist.ModelList.ModelName;
            newRQ.Remark = Remark;

            //2. Tạo số DTF
            newRQ.DTFNo = genarateDTFNo("", true, false);
            //3. update progress and next status
            newRQ = genarateProgress(newRQ);

            // Lưu
            db.DieLendingRequests.Add(newRQ);
            db.SaveChanges();
            sendEmailJob.sendEmailDieTransfer(newRQ);
            status = true;
            msg = "OK";
        exit:
            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getRequest(int? id)
        {
            db.Configuration.ProxyCreationEnabled = false;
            var output = db.DieLendingRequests.Find(id);
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult verifyDTF(string[] id, string action, string remark, string actualTransferDate, string reason)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Lending_Role"].ToString();

            List<string> success = new List<string>();
            List<string> fail = new List<string>();

            var today = DateTime.Now;
            for (var i = 0; i < id.Length; i++)
            {
                var dtf = db.DieLendingRequests.Find(int.Parse(id[i]));

                if (action.ToUpper().Contains("CONFIRM") && dept.Contains(dtf.RequestDept) && dtf.LendingStatusID == 9)
                {
                    dtf = genarateProgress(dtf);
                    dtf.ActualTransferDate = DateTime.Parse(actualTransferDate);
                    // Update new location for die
                    Die1 die = db.Die1.Where(x => x.DieNo == dtf.DieNo && x.Active != false && x.isCancel != true).FirstOrDefault();
                    die.SupplierID = (int)dtf.NewLocationID;
                    db.Entry(die).State = EntityState.Modified;
                    db.Entry(dtf).State = EntityState.Modified;
                    db.SaveChanges();
                    success.Add(dtf.DTFNo);
                    goto Exit;
                }

                if ((dtf.LendingStatusCategory.RoleRespone == role || role == "Approve") && dept.Contains(dtf.LendingStatusCategory.DeptRespone))
                {
                    if (action.ToUpper().Contains("CHECK") || action.ToUpper().Contains("APPROVE"))
                    {
                        dtf = genarateProgress(dtf);
                        dtf.ControlNo = !String.IsNullOrWhiteSpace(remark) ? remark + System.Environment.NewLine + dtf.ControlNo : dtf.ControlNo;
                    }
                    if (action.ToUpper().Contains("REJECT") && dept.Contains(dtf.LendingStatusCategory.DeptRespone))
                    {
                        if (String.IsNullOrWhiteSpace(reason))
                        {
                            goto Exit;
                        }
                        dtf.LendingStatusID = 11;
                        dtf.ControlNo = today.ToString("yyyy/MM/dd_") + Session["Name"].ToString() + " Reject RQ: " + reason + System.Environment.NewLine + dtf.ControlNo;
                    }

                    if (action.ToUpper().Contains("CANCEL") && dept.Contains(dtf.LendingStatusCategory.DeptRespone))
                    {
                        if (String.IsNullOrWhiteSpace(reason))
                        {
                            goto Exit;
                        }
                        dtf.LendingStatusID = 12;
                        dtf.ControlNo = today.ToString("yyyy/MM/dd_") + Session["Name"].ToString() + " Cancel RQ: " + reason + System.Environment.NewLine + dtf.ControlNo;
                    }

                    db.Entry(dtf).State = EntityState.Modified;
                    db.SaveChanges();
                    success.Add(dtf.DTFNo);

                    sendEmailJob.sendEmailDieTransfer(dtf);
                }
                else
                {
                    fail.Add(dtf.DTFNo + "_ You have not permit or not your turn!");
                }

            Exit:
                ViewBag.forExit = "Just for exit loop";


            }



            return Json(new { success = success, fail = fail }, JsonRequestBehavior.AllowGet);
        }


        public ActionResult ExportToForm(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var dtf = db.DieLendingRequests.Find(id);
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/DieTransfer/Format/FormDTF.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();

                sheet.Cells["M4"].Value = dtf.DTFNo;
                sheet.Cells["V4"].Value = dtf.ModelName;
                sheet.Cells["G5"].Value = dtf.LendingTypeCategory.Type;
                sheet.Cells["G6"].Value = dtf.DieNo;
                sheet.Cells["G7"].Value = dtf.FixedAssetNo;
                sheet.Cells["Q5"].Value = dtf.CurrentLocation;
                sheet.Cells["Q6"].Value = dtf.NewLocation;
                sheet.Cells["Q7"].Value = dtf.ActualShot;
                sheet.Cells["Y5"].Value = dtf.ETDPlan.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.ETDPlan) : "-";
                sheet.Cells["Y6"].Value = dtf.ETAPlan.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.ETAPlan) : "-";
                sheet.Cells["Y7"].Value = dtf.Transport;
                sheet.Cells["B9"].Value = dtf.Remark;


                if (dtf.LendingStatusID != 11 && dtf.LendingStatusID != 12)
                {

                    if (dtf.LendingTypeID == 1)
                    {
                        sheet.Cells["B12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells["B12"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("red"));
                        sheet.Cells["C14"].Value = dtf.Requestor;
                        sheet.Cells["C16"].Value = dtf.RequestDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.RequestDate) : "-";
                        sheet.Cells["G14"].Value = dtf.DMTCheckBy;
                        sheet.Cells["G16"].Value = dtf.DMTCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTCheckDate) : "-";
                        sheet.Cells["K14"].Value = dtf.DMTAppBy;
                        sheet.Cells["K16"].Value = dtf.DMTAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTAppDate) : "-";
                    }
                    if (dtf.LendingTypeID == 2)
                    {
                        sheet.Cells["B18"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells["B18"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("red"));
                        sheet.Cells["C20"].Value = dtf.Requestor;
                        sheet.Cells["C22"].Value = dtf.RequestDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.RequestDate) : "-";
                        sheet.Cells["G20"].Value = dtf.DMTCheckBy;
                        sheet.Cells["G22"].Value = dtf.DMTCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTCheckDate) : "-";
                        sheet.Cells["K20"].Value = dtf.DMTAppBy;
                        sheet.Cells["K22"].Value = dtf.DMTAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTAppDate) : "-";

                        sheet.Cells["O20"].Value = dtf.PURCheckBy;
                        sheet.Cells["O22"].Value = dtf.PURCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURCheckDate) : "-";
                        sheet.Cells["R20"].Value = dtf.PURAppBy;
                        sheet.Cells["R22"].Value = dtf.PURAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURAppDate) : "-";

                        sheet.Cells["U20"].Value = dtf.PUCCheckBy;
                        sheet.Cells["U22"].Value = dtf.PUCCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCCheckDate) : "-";
                        sheet.Cells["X20"].Value = dtf.PUCAppBy;
                        sheet.Cells["X22"].Value = dtf.PUCAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCAppDate) : "-";
                    }
                    if (dtf.LendingTypeID == 3)
                    {
                        sheet.Cells["B24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells["B24"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("red"));
                        sheet.Cells["C26"].Value = dtf.Requestor;
                        sheet.Cells["C28"].Value = dtf.RequestDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.RequestDate) : "-";
                        sheet.Cells["G26"].Value = dtf.PURCheckBy;
                        sheet.Cells["G28"].Value = dtf.PURCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURCheckDate) : "-";
                        sheet.Cells["K26"].Value = dtf.PURAppBy;
                        sheet.Cells["K28"].Value = dtf.PURAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURAppDate) : "-";

                        sheet.Cells["P26"].Value = dtf.PUCCheckBy;
                        sheet.Cells["P28"].Value = dtf.PUCCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCCheckDate) : "-";
                        sheet.Cells["U26"].Value = dtf.PUCAppBy;
                        sheet.Cells["U28"].Value = dtf.PUCAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCAppDate) : "-";
                    }
                    if (dtf.LendingTypeID == 4)
                    {
                        sheet.Cells["B30"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells["B30"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("red"));
                        sheet.Cells["C32"].Value = dtf.Requestor;
                        sheet.Cells["C34"].Value = dtf.RequestDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.RequestDate) : "-";
                        sheet.Cells["G32"].Value = dtf.PURCheckBy;
                        sheet.Cells["G34"].Value = dtf.PURCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURCheckDate) : "-";
                        sheet.Cells["K32"].Value = dtf.PURAppBy;
                        sheet.Cells["K34"].Value = dtf.PURAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PURAppDate) : "-";

                        string[] configCodeInhouse = { "5400", "5500", "3400", "3500" };
                        if (Array.IndexOf(configCodeInhouse, dtf.NewLocation) != -1) // Khuon chuyen ve inhouse
                        {
                            sheet.Cells["O32"].Value = dtf.DMTCheckBy;
                            sheet.Cells["O34"].Value = dtf.DMTCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTCheckDate) : "-";
                            sheet.Cells["R32"].Value = dtf.DMTAppBy;
                            sheet.Cells["R34"].Value = dtf.DMTAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.DMTAppDate) : "-";

                        }
                        else // Khuon chuyen tu diemake => Supplier
                        {
                            sheet.Cells["O32"].Value = dtf.PAECheckBy;
                            sheet.Cells["O34"].Value = dtf.PAECheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PAECheckDate) : "-";
                            sheet.Cells["R32"].Value = dtf.PAEAppBy;
                            sheet.Cells["R34"].Value = dtf.PAEAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PAEAppDate) : "-";

                        }
                        sheet.Cells["U32"].Value = dtf.PUCCheckBy;
                        sheet.Cells["U34"].Value = dtf.PUCCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCCheckDate) : "-";
                        sheet.Cells["X32"].Value = dtf.PUCAppBy;
                        sheet.Cells["X34"].Value = dtf.PUCAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCAppDate) : "-";
                    }
                    if (dtf.LendingTypeID == 5)
                    {
                        sheet.Cells["B36"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells["B36"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("red"));
                        sheet.Cells["C38"].Value = dtf.Requestor;
                        sheet.Cells["C40"].Value = dtf.RequestDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.RequestDate) : "-";
                        sheet.Cells["G38"].Value = dtf.PUCCheckBy;
                        sheet.Cells["G40"].Value = dtf.PUCCheckDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCCheckDate) : "-";
                        sheet.Cells["K38"].Value = dtf.PUCAppBy;
                        sheet.Cells["K40"].Value = dtf.PUCAppDate.HasValue ? String.Format("{0:yyyy-MM-dd}", dtf.PUCAppDate) : "-";
                    }
                }

                sheet.Protection.IsProtected = true;
                sheet.Protection.SetPassword("DMSPROTECTION");
                package.SaveAs(output);
            }
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=" + dtf.DTFNo + ".xlsx");

            output.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();
            return RedirectToAction("Index");
        }

       
        public ActionResult exportToList(List<DieLendingRequest> ListRQ)
        {
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/DieTransfer/Format/FormExportListDTF.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();
                int rowId = 4;
                int i = 1;
                foreach (var rq in ListRQ)
                {
                    sheet.Cells["A" + rowId.ToString()].Value = i;
                    sheet.Cells["B" + rowId.ToString()].Value = rq.DTFNo;
                    sheet.Cells["C" + rowId.ToString()].Value = rq.LendingStatusCategory.Type;
                    sheet.Cells["D" + rowId.ToString()].Value = rq.LendingTypeCategory.Type;
                    sheet.Cells["E" + rowId.ToString()].Value = rq.DieNo;
                    sheet.Cells["F" + rowId.ToString()].Value = rq.FixedAssetNo;
                    sheet.Cells["G" + rowId.ToString()].Value = rq.CurrentLocation;
                    sheet.Cells["H" + rowId.ToString()].Value = rq.NewLocation;
                    sheet.Cells["I" + rowId.ToString()].Value = rq.ModelName;
                    sheet.Cells["J" + rowId.ToString()].Value = rq.ActualShot;
                    sheet.Cells["K" + rowId.ToString()].Value = rq.Transport;
                    sheet.Cells["L" + rowId.ToString()].Value = rq.ETDPlan;
                    sheet.Cells["M" + rowId.ToString()].Value = rq.ETAPlan;
                    sheet.Cells["N" + rowId.ToString()].Value = rq.ActualTransferDate;
                    sheet.Cells["O" + rowId.ToString()].Value = rq.RequestDept;
                    sheet.Cells["P" + rowId.ToString()].Value = rq.Requestor;
                    sheet.Cells["Q" + rowId.ToString()].Value = rq.RequestDate;
                    sheet.Cells["R" + rowId.ToString()].Value = rq.DMTCheckBy;
                    sheet.Cells["S" + rowId.ToString()].Value = rq.DMTCheckDate;
                    sheet.Cells["T" + rowId.ToString()].Value = rq.DMTAppBy;
                    sheet.Cells["U" + rowId.ToString()].Value = rq.DMTAppDate;
                    sheet.Cells["V" + rowId.ToString()].Value = rq.PURCheckBy;
                    sheet.Cells["W" + rowId.ToString()].Value = rq.PURCheckDate;
                    sheet.Cells["X" + rowId.ToString()].Value = rq.PURAppBy;
                    sheet.Cells["Y" + rowId.ToString()].Value = rq.PURAppDate;
                    sheet.Cells["Z" + rowId.ToString()].Value = rq.PUCCheckBy;
                    sheet.Cells["AA" + rowId.ToString()].Value = rq.PUCCheckDate;
                    sheet.Cells["AB" + rowId.ToString()].Value = rq.PUCAppBy;
                    sheet.Cells["AC" + rowId.ToString()].Value = rq.PUCAppDate;
                    sheet.Cells["AD" + rowId.ToString()].Value = rq.PAECheckBy;
                    sheet.Cells["AE" + rowId.ToString()].Value = rq.PAECheckDate;
                    sheet.Cells["AF" + rowId.ToString()].Value = rq.PAEAppBy;
                    sheet.Cells["AG" + rowId.ToString()].Value = rq.PAEAppDate;
                    sheet.Cells["AH" + rowId.ToString()].Value = rq.Remark;
                    i++;
                    rowId++;
                }

                package.SaveAs(output);
            }
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=Die_Transfer_List" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx");
            output.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();
            return RedirectToAction("Index");
        }






        public JsonResult deleteRequest(int id)
        {
            var status = false;
            var msg = "";
            var dept = Session["Dept"].ToString();
            var role = Session["Lending_Role"].ToString();
            var rq = db.DieLendingRequests.Find(id);

            if ((role == "Check" || role == "Approve") && dept == rq.RequestDept && rq.LendingStatusCategory.IsCanIssue == true)
            {
                rq.LendingStatusID = 9;
                rq.Active = false;

                db.Entry(rq).State = EntityState.Modified;
                db.SaveChanges();
                status = true;
            }
            else
            {
                if ((role != "Check" || role != "Approve"))
                {
                    msg = "You do not have permision to delete it!";
                }
                else
                {
                    if (dept != rq.RequestDept)
                    {
                        msg = "You(" + dept + ") can not delete request of " + rq.RequestDept + " dept!";
                    }
                    else
                    {
                        msg = "This request is in processing by other dept, You can not delete it";
                    }
                }
                status = false;
            }
            var data = new
            {
                status = status,
                msg = msg,
                id = id
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }

       



        public string genarateDTFNo(string currentDTFNo, bool isNew, bool isRevise)
        {
            var output = "";
            var today = DateTime.Now;
            var totalDTFinThisYear = db.DieLendingRequests.Where(x => x.RequestDate.Value.Year == today.Year).Count() + 1;
            if (isNew)
            {
                output = "DTF" + today.ToString("yyMMdd-") + totalDTFinThisYear + "-00";
            }
            if (isRevise)
            {
                var upver = currentDTFNo.Substring(currentDTFNo.Length - 2, 2); // 00
                int upverInt = Convert.ToInt16(upver) + 1;
                string upverStr = Convert.ToString(upverInt);
                if (upverStr.Length == 1)
                {
                    upverStr = "0" + upverStr;
                }
                var mainNo = currentDTFNo.Remove(currentDTFNo.Length - 2, 2);
                output = mainNo + upverStr;
            }
            return output;
        }

        public DieLendingRequest genarateProgress(DieLendingRequest rq)
        {
            var route = rq.LendingTypeID;
            var currentStatusID = rq.LendingStatusID;
            var newStatus = 0;
            int[] configRoute1Status = { 1, 2, 9, 10 };
            int[] configRoute2Status = { 1, 2, 3, 4, 5, 6, 9, 10 };
            int[] configRoute3Status = { 3, 4, 5, 6, 9, 10 };
            int[] configRoute4Status_Inhouse = { 3, 4, 1, 2, 5, 6, 9, 10 };
            int[] configRoute4Status_Supplier = { 3, 4, 7, 8, 5, 6, 9, 10 };
            int[] configRoute5Status = { 5, 6, 9, 10 };
            var today = DateTime.Now;
            // Route = 1~5
            if (route == 1) // status: 1 => 2 => 1(recive) => 2(reciept)
            {
                newStatus = currentStatusID == null ? configRoute1Status[0] : configRoute1Status[Array.IndexOf(configRoute1Status, currentStatusID) + 1];
            }
            if (route == 2)
            {
                newStatus = currentStatusID == null ? configRoute2Status[0] : configRoute2Status[Array.IndexOf(configRoute2Status, currentStatusID) + 1];
            }
            if (route == 3)
            {
                newStatus = currentStatusID == null ? configRoute3Status[0] : configRoute3Status[Array.IndexOf(configRoute3Status, currentStatusID) + 1];
            }

            if (route == 4)
            {
                string[] configCodeInhouse = { "5400", "5500", "3400", "3500" };
                if (Array.IndexOf(configCodeInhouse, rq.NewLocation) != -1) // Khuon chuyen ve inhouse
                {
                    newStatus = currentStatusID == null ? configRoute4Status_Inhouse[0] : configRoute4Status_Inhouse[Array.IndexOf(configRoute4Status_Inhouse, currentStatusID) + 1];

                }
                else // Khuon chuyen tu diemake => Supplier
                {
                    newStatus = currentStatusID == null ? configRoute4Status_Supplier[0] : configRoute4Status_Supplier[Array.IndexOf(configRoute4Status_Supplier, currentStatusID) + 1];

                }
            }
            if (route == 5)
            {
                newStatus = currentStatusID == null ? configRoute5Status[0] : configRoute5Status[Array.IndexOf(configRoute5Status, currentStatusID) + 1];
            }

            // Who
            if (currentStatusID == null)
            {
                rq.Requestor = Session["Name"].ToString();
                rq.RequestDate = today;
                rq.RequestDept = Session["Dept"].ToString();
                rq.Active = true;
            }
            if (currentStatusID == 1)
            {
                rq.DMTCheckBy = Session["Name"].ToString();
                rq.DMTCheckDate = today;
            }
            if (currentStatusID == 2)
            {
                rq.DMTAppBy = Session["Name"].ToString();
                rq.DMTAppDate = today;
            }
            if (currentStatusID == 3)
            {
                rq.PURCheckBy = Session["Name"].ToString();
                rq.PURCheckDate = today;
            }
            if (currentStatusID == 4)
            {
                rq.PURAppBy = Session["Name"].ToString();
                rq.PURAppDate = today;
            }
            if (currentStatusID == 5)
            {
                rq.PUCCheckBy = Session["Name"].ToString();
                rq.PUCCheckDate = today;
            }
            if (currentStatusID == 6)
            {
                rq.PUCAppBy = Session["Name"].ToString();
                rq.PUCAppDate = today;
            }
            if (currentStatusID == 7)
            {
                rq.PAECheckBy = Session["Name"].ToString();
                rq.PAECheckDate = today;
            }
            if (currentStatusID == 8)
            {
                rq.PAEAppBy = Session["Name"].ToString();
                rq.PAEAppDate = today;
            }

            rq.LendingStatusID = newStatus;
            return rq;
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
