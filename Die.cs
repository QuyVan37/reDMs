using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects.DataClasses;
using System.Deployment.Internal;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Windows.Media.Media3D;
using Antlr.Runtime;
using Aspose.Slides;
using DMS03.Models;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Org.BouncyCastle.Crypto;
using PagedList;
//using static DMS03.Controllers.Die_Launch_ManagementController;
using static iTextSharp.text.pdf.AcroFields;

namespace DMS03.Controllers
{
    public class DiesController : Controller
    {
        private DMSEntities db = new DMSEntities();
        private ECN_SystemEntities ecn = new ECN_SystemEntities();
        public CommonFunctionController commonFunction = new CommonFunctionController();
        public SendEmailController mailJob = new SendEmailController();
       

        //public async void LinkECN()
        //{
        //    try
        //    {
        //        var data = new { PartNo = "RC5-2015-000"};

        //        string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(data);

        //        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
        //        HttpResponseMessage response = await HttpClient.PostAsync("http://prodqv:7200/api/PQA/proc_insert_from_awi_to_ecn", content);
        //        if (response.IsSuccessStatusCode)
        //        {
        //            string apiResponse = await response.Content.ReadAsStringAsync();
        //            ViewBag.Msg = apiResponse.ToString();
        //        }
        //        else
        //        {
        //            ViewBag.Msg = "Failed to call API";
        //        }
        //        ViewBag.Msg = "Update data successfully.";
        //    }
        //    catch (HttpRequestException ex)
        //    {
        //        ViewBag.Msg = "Failed to update data";
        //    }

        //}



        // GET: Dies
        public ActionResult Index(string Search, string partName, string dieNo, int? page, string export)
        {
            if (page == null) page = 1;
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            if (Session["Role"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            List<CommonDie1> searchResult = new List<CommonDie1>();
            if (!String.IsNullOrEmpty(Search))
            {
                //Search PartNo
                searchResult = db.CommonDie1.Where(x => x.PartNo.Contains(Search.Trim())).ToList();
                //Search Die No
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.DieNo.Contains(Search.Trim())).ToList();
                }
                //Search Part Name
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.Parts1.PartName.Contains(Search.Trim())).ToList();
                }
                //Search Supplier Name
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.Die1.Supplier.SupplierName.Contains(Search.Trim())).ToList();
                }
                //Search Supplier Code
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.Die1.Supplier.SupplierCode.Contains(Search.Trim())).ToList();
                }
                //Search Material
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.Parts1.Material.Contains(Search.Trim())).ToList();
                }
                //Search Model
                if (searchResult.Count() == 0)
                {
                    searchResult = db.CommonDie1.Where(x => x.Parts1.Model.Contains(Search.Trim())).ToList();
                }

            }
            else
            {
                searchResult = db.CommonDie1.Include(c => c.Die1).Include(c => c.Parts1).OrderByDescending(x => x.DieID).Take(10).ToList();
            }
            if (!String.IsNullOrEmpty(dieNo))
            {
                searchResult = db.CommonDie1.Where(x => x.DieNo == dieNo.Trim()).ToList();
            }

            if (export == "Export")
            {
                if (String.IsNullOrEmpty(Search))
                {
                    searchResult = db.CommonDie1.Include(c => c.Die1).Include(c => c.Parts1).OrderByDescending(x => x.DieID).ToList();
                }

                // ExportExcel(searchResult);
            }

            ViewBag.search = Search;
            return View(searchResult.Where(x => x.Active != false).ToPagedList(pageNumber, pageSize));
        }

        // GET: Dies/PartInfor/id
        public ActionResult PartInfor(int? id, string partNo, int? page)
        {
            if (page == null) page = 1;
            int pageSize = 100;
            int pageNumber = (page ?? 1);
            List<ECN_PART_COVER> part = new List<ECN_PART_COVER>();

            if (!String.IsNullOrWhiteSpace(partNo))
            {
                part = ecn.ECN_PART_COVER.Where(x => x.Part_No.Contains(partNo)).OrderByDescending(x => x.Date_Entry).ToList();
            }

            return View(part.ToPagedList(pageNumber, pageSize));
        }








        public ActionResult index2()
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }

            ViewBag.FA_Result = new SelectList(commonFunction.faResultCategory(), "value", "show");
            ViewBag.clasify = new SelectList(clasify_list(), "value", "show");
            ViewBag.status = new SelectList(statusCategory(), "value", "show");
            ViewBag.deptResponse = new SelectList(deptResponse(), "value", "show");
            ViewBag.pendingStatusCategory = new SelectList(db.DieLaunchingWarningConfigs, "Status", "Status");
            ViewBag.ModelID = new SelectList(db.ModelLists.Where(x => x.Active != false), "ModelID", "ModelName");
            ViewBag.ProcessCode = new SelectList(db.ProcessCodeCalogories, "ProcessCodeID", "Type");
            ViewBag.SupplierID = new SelectList(db.Suppliers.Where(x => x.Active != false).Select(x => new { SupplierID = x.SupplierID, SupplierCode = x.SupplierCode + "-" + x.SupplierName }), "SupplierID", "SupplierCode");
            return View();
        }

        public JsonResult updateCategory()
        {
            var allDie = db.Die1.Where(x => x.Active != false && x.isCancel != true);
            var output = new
            {
                MP = allDie.Where(x => x.DieClassify.Contains("MP")).Count(),
                MT = allDie.Where(x => x.DieClassify.Contains("MT")).Count(),
                Renewal = allDie.Where(x => x.DieClassify.Contains("Renewal")).Count(),
                Additional = allDie.Where(x => x.DieClassify.Contains("Additional")).Count(),
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getData(int? dieID, string following, string search, string[] modelID, string[] supplierID, string[] category, string[] FAresult, string[] procesCode,
    string POFrom, string POTo, string[] status, string[] statusPending, string deptRes, string export, string supplierCode, int? page)
        {
            // db.Configuration.ProxyCreationEnabled = false;
            int pageIndex = page ?? 1;
            int pageSize = 10;
            //List<CommonDie1> records = db.CommonDie1.ToList();
            //var records = db.CommonDie1.Where(x => x.Active != false).Include(x => x.Die1).Include(x => x.Parts1).ToList();
            modelID = modelID == null ? new string[0] { } : modelID.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            supplierID = supplierID == null ? new string[0] { } : supplierID.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            category = category == null ? new string[0] { } : category.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            FAresult = FAresult == null ? new string[0] { } : FAresult.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            procesCode = procesCode == null ? new string[0] { } : procesCode.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            status = status == null ? new string[0] { } : status.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            statusPending = statusPending == null ? new string[0] { } : statusPending.Where(x => !String.IsNullOrEmpty(x)).ToArray();

            if (!String.IsNullOrWhiteSpace(supplierCode))
            {
                string id = db.Suppliers.Where(x => x.SupplierCode == supplierCode).FirstOrDefault()?.SupplierID.ToString();
                supplierID = supplierID.Append(id).ToArray();
            }

            List<CommonDie1> records = new List<CommonDie1>();
            if (dieID > 0)
            {
                records = db.CommonDie1.Where(x => x.DieID == (int)dieID).ToList();
                goto exit;
            }

            if (!String.IsNullOrWhiteSpace(search))
            {
                search = search.Trim().ToUpper();
                // cần search toàn bộ common, Family die
                records = db.CommonDie1.Where(x => x.Active != false && x.PartNo.Contains(search)).ToList();
                if (records.Count() == 0)
                {
                    records = db.CommonDie1.Where(x => x.Active != false && x.DieNo.Contains(search)).ToList();
                }
                goto exit;
            }



            if (modelID.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var id in modelID)
                {
                    int intID = int.Parse(id);
                    var res = db.CommonDie1.Where(x => x.Die1.ModelID == intID).ToList();
                    searchResult.AddRange(res);
                }
                records = searchResult;
                if (records.Count() == 0) goto exit;

            }

            if (supplierID.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var id in supplierID)
                {
                    int intID = int.Parse(id);
                    if (records.Count() == 0)
                    {
                        var res = db.CommonDie1.Where(x => x.Die1.SupplierID == intID).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => x.Die1.SupplierID == intID).ToList();
                        searchResult.AddRange(res);
                    }

                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            if (category.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var item in category)
                {
                    if (records.Count() == 0)
                    {
                        var res = db.CommonDie1.Where(x => (x.Die1.DieClassify).Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => (x.Die1.DieClassify).Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }

                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            if (FAresult.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var item in FAresult)
                {
                    if (records.Count() == 0)
                    {
                        var res = db.CommonDie1.Where(x => x.Die1.FA_Result != null).Where(x => x.Die1.FA_Result.Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => x.Die1.FA_Result != null).Where(x => x.Die1.FA_Result.Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }
                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            if (procesCode.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var item in procesCode)
                {
                    int processCodeID = int.Parse(item);
                    if (records.Count() == 0)
                    {
                        var res = db.CommonDie1.Where(x => x.Die1.ProcessCodeID == processCodeID).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => x.Die1.ProcessCodeID == int.Parse(item)).ToList();
                        searchResult.AddRange(res);
                    }

                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            //*****************************

            if (status.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var item in status)
                {
                    if (records.Count() == 0)
                    {
                        var allDie = db.CommonDie1.ToList();
                        var res = allDie.Where(x => !x.Die1.DieClassify.Contains("MP")).Where(x => commonFunction.getStatus(x.Die1).Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => commonFunction.getStatus(x.Die1).Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }

                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            if (statusPending.Length > 0)
            {
                List<CommonDie1> searchResult = new List<CommonDie1>();
                foreach (var item in statusPending)
                {
                    if (records.Count() == 0)
                    {
                        var allDie = db.CommonDie1.ToList();
                        var res = allDie.Where(x => !x.Die1.DieClassify.Contains("MP")).Where(x => commonFunction.getPendingAndDeptResponse(x.Die1).Pending_Status.Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        var res = records.Where(x => commonFunction.getPendingAndDeptResponse(x.Die1).Pending_Status.Contains(item)).ToList();
                        searchResult.AddRange(res);
                    }

                }
                records = searchResult;
                if (records.Count() == 0) goto exit;
            }

            if (!string.IsNullOrEmpty(deptRes))
            {
                if (records.Count() == 0)
                {
                    var allDie = db.CommonDie1.ToList();
                    records = allDie.Where(x => !x.Die1.DieClassify.Contains("MP")).Where(x => commonFunction.getPendingAndDeptResponse(x.Die1).Dept_Respone.Contains(deptRes)).ToList();
                }
                else
                {
                    records = records.Where(x => commonFunction.getPendingAndDeptResponse(x.Die1).Dept_Respone.Contains(deptRes)).ToList();
                    if (records.Count() == 0) goto exit;
                }
            }

            //********************************************************

            if (!string.IsNullOrEmpty(POFrom))
            {
                if (records.Count() == 0)
                {
                    records = db.CommonDie1.Where(x => x.Die1.PODate != null)
              .Where(x => x.Die1.PODate >= DateTime.Parse(POFrom)).ToList();
                }
                else
                {
                    records = records.Where(x => x.Die1.PODate != null)
           .Where(x => x.Die1.PODate >= DateTime.Parse(POFrom)).ToList();

                }
                if (records.Count() == 0) goto exit;
            }
            if (!string.IsNullOrEmpty(POTo))
            {
                if (records.Count() == 0)
                {
                    records = db.CommonDie1.Where(x => x.Die1.PODate != null)
               .Where(x => x.Die1.PODate <= DateTime.Parse(POTo)).ToList();
                }
                else
                {
                    records = records.Where(x => x.Die1.PODate != null)
           .Where(x => x.Die1.PODate <= DateTime.Parse(POTo)).ToList();
                }

                if (records.Count() == 0) goto exit;
            }

        exit:

            records = records.Where(x => x.Active != false && x.Die1.Active != false).ToList();
            Object exportAddress;
            if (export == "Export")
            {
                exportAddress = exportToControlList(records);
                return Json(exportAddress, JsonRequestBehavior.AllowGet);
            }

            var result = dataReturnView((List<CommonDie1>)records.Page(pageIndex, pageSize).ToList());
            double totalPage = decimal.ToDouble(records.Count()) / decimal.ToDouble(pageSize);
            var output = new
            {
                test = totalPage,
                page = pageIndex,
                totalPage = Math.Ceiling(totalPage),
                data = result,
            };

            return Json(output, JsonRequestBehavior.AllowGet);
        }



        public JsonResult AddNew(string PartNo, string DieCode, string ProcessCode, string SupplierID, string ModelID, string needUseDate, string targetOK)
        {
            var dept = Session["Dept"].ToString();
            string[] cfDeptCanAddNew = { "CRG", "PAE", "PUR" };

            bool status = false;
            string msg = "";
            if (Array.IndexOf(cfDeptCanAddNew, dept) == -1)
            {
                status = false;
                msg = "You do not have permition!";
                goto exit;
            }
            // Check du lieu use nhap vao
            if (!String.IsNullOrWhiteSpace(PartNo) && !String.IsNullOrWhiteSpace(DieCode))
            {
                if (PartNo.Trim().ToUpper().Length != 12)
                {
                    status = false;
                    msg = "Part No MUST 12 letter like XXX-XXXX-000";
                    goto exit;
                }
                if (DieCode.Trim().ToUpper().Length != 3)
                {
                    status = false;
                    msg = "Part No MUST 3 letter such as 11A, 21A, 14A,...";
                    goto exit;
                }
                if (String.IsNullOrWhiteSpace(ProcessCode))
                {
                    status = false;
                    msg = "Pls chose process code MO/PX...!";
                    goto exit;
                }
                if (String.IsNullOrWhiteSpace(ModelID))
                {
                    status = false;
                    msg = "Pls chose model Name!";
                    goto exit;
                }
                if (String.IsNullOrWhiteSpace(SupplierID))
                {
                    SupplierID = "0";
                }
                string dieNo = PartNo.Trim().ToUpper() + "-" + DieCode.Trim().ToUpper() + "-001";
                var checkDieResult = commonFunction.checkDieExist(dieNo);
                if (checkDieResult.isExist == true)
                {
                    status = false;
                    msg = "This Die already Exist!, You can not Add New";
                    goto exit;
                }
                else
                {
                    // add new die to database
                    var dieID = commonFunction.genarateNewDie(dieNo, DieCode, PartNo, ProcessCode, ModelID, SupplierID, needUseDate, targetOK, dept, null, 0);
                    if (dieID > 0)
                    {
                        status = true;
                        msg = "Success!";
                        //mailJob.anounceNewDie(PartNo, DieCode, db.ModelLists.Find(int.Parse(ModelID)).ModelName, commonFunction.isRenewOrAddOrMT(DieCode), "Pls PUS Select supplier for this die!");

                    }
                    else
                    {
                        status = false;
                        msg = "Error, Can not add to DB!";
                    }
                }
            }
            else
            {
                status = false;
                msg = "Pls input Part No and Die No";
                goto exit;
            }


        exit:
            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult UploadListNewDie(HttpPostedFileBase file)
        {
            string err = "";
            if (file == null)
            {
                return Json(new { status = false, msg = "No file input!" }, JsonRequestBehavior.AllowGet);
            }

            string handle = Guid.NewGuid().ToString();
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(file.InputStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                for (int row = start.Row + 3; row <= end.Row; row++)
                { // Row by row...
                  //Hãy kiểm tra điều kiện đã có dữ liệu nhập vào excel file
                    var PartNo = worksheet.Cells[row, 2].Text.Trim().ToUpper();
                    if (String.IsNullOrEmpty(PartNo)) break;
                    if (PartNo.Length != 12)
                    {
                        err = "Error: Part No gồm 12 kí tự  XXX-XXXX-000";
                        goto exitLoop;
                    }
                    var DieNo = worksheet.Cells[row, 3].Text.Trim().ToUpper();
                    if (DieNo.Length != 3)
                    {
                        err = "Error: Die No gồm 3 kí tự such as 11A, 14A, 21A,...";
                        goto exitLoop;
                    }

                    var ProcessCode = worksheet.Cells[row, 4].Text.ToUpper();
                    var findProcessCode = db.ProcessCodeCalogories.Where(x => x.Type == ProcessCode).FirstOrDefault();
                    if (findProcessCode == null)
                    {
                        err = "Error: ProcessCode ko input hoặc tồn tại ";
                        goto exitLoop;
                    }

                    var modelName = worksheet.Cells[row, 5].Text.Trim().ToUpper();
                    if (modelName.Length == 0)
                    {
                        err = "Error:Không input model ";
                        goto exitLoop;
                    }

                    string SUPPLIERID = "";
                    var SupplierCode = worksheet.Cells[row, 6].Text.Trim().ToUpper();
                    var FindSupplier = db.Suppliers.Where(x => x.SupplierCode == SupplierCode).FirstOrDefault();
                    if (FindSupplier == null)
                    {
                        SUPPLIERID = "0";
                    }
                    else
                    {
                        SUPPLIERID = FindSupplier.SupplierID.ToString();
                    }

                    var needUseDate = worksheet.Cells[row, 7].Text.Trim().ToUpper();
                    var targetOK = worksheet.Cells[row, 8].Text.Trim().ToUpper();
                    // Xử lí dữ liệu
                    var model = commonFunction.CreateModelName(modelName);
                    err = AddNew(PartNo, DieNo, findProcessCode.ProcessCodeID.ToString(), SUPPLIERID, model[0], needUseDate, targetOK).Data.ToString();


                exitLoop:
                    ViewBag.note = "Câu lệnh vô nghĩ để thoát khỏi vòng lặp";
                    worksheet.Cells[row, 9].Value = err;
                }
                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }

            var data = new { status = true, FileGuid = handle, FileName = "ResultUploadAddNewDieToDMS.xlsx" };

            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CloseItem(string[] ids, int? ModelID)
        {
            // close sau khi có first lost cho trường hợp RN & AD
            // close cho MT model sau khi PIC close Model
            ids = ids == null ? new string[0] { } : ids.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            var dept = Session["Dept"].ToString();
            var role = Session["Admin"].ToString();
            int inPut = ids.Length;
            int outPut = 0;
            if (dept != "PAE" && role != "Admin")
            {
                return Json(new { status = false, msg = "You do not have permition!" }, JsonRequestBehavior.AllowGet);
            }

            if (ids.Length > 0)
            {
                foreach (var id in ids)
                {

                    Die1 exitDie = db.Die1.Find(int.Parse(id));
                    if ((exitDie.isCancel != true && !String.IsNullOrWhiteSpace(exitDie.First_Lot_Date)) || exitDie.Belong != "LBP")
                    {
                        exitDie.isClosed = true;
                        exitDie.DieClassify = "MP";
                        exitDie.DieStatusID = 3; // MP_Main
                        // Stop renewal die
                        if (commonFunction.isRenewOrAddOrMT(exitDie.Die_Code) == "Renewal" && exitDie.Belong == "LBP")
                        {
                            var orignalDie = commonFunction.getOriginalDie(exitDie);
                            if (orignalDie != null)
                            {
                                orignalDie.DieStatusID = 8; // Stop
                                orignalDie.StopDate = DateTime.Now;
                                orignalDie.RemarkDieStatusUsing = "Renewal " + exitDie.DieNo + "has first lot on " + exitDie.First_Lot_Date;
                                db.Entry(orignalDie).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                        }
                        // exitDie = commonFunction.updateDieStatus(exitDie);
                        db.Entry(exitDie).State = EntityState.Modified;
                        db.SaveChanges();
                        outPut = outPut + 1;
                    }

                }

            }
            if (ModelID != null)
            {
                List<Die1> Dies = db.Die1.Where(x => x.Active != false && x.isCancel != true && x.DieClassify.Contains("MT") && x.ModelID == ModelID).ToList();
                inPut = Dies.Count();
                foreach (var die in Dies)
                {
                    Die1 exitDie = db.Die1.Find(die.DieID);
                    if (exitDie.isCancel != true)
                    {
                        exitDie.isClosed = true;
                        exitDie.DieClassify = "MP";
                        exitDie.DieStatusID = 3; // MP_Main
                        //exitDie = commonFunction.updateDieStatus(exitDie);
                        db.Entry(exitDie).State = EntityState.Modified;
                        db.SaveChanges();
                        outPut = outPut + 1;
                    }

                }
                var model = db.ModelLists.Find(ModelID);
                model.Phase = "MP";
                model.MPDate = DateTime.Now;
                db.Entry(model).State = EntityState.Modified;
                db.SaveChanges();
            }

            return Json(new { inPut = inPut, success = outPut }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CancelItem(string[] ids)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Die_Lauch_Role"].ToString();
            int inPut = ids.Length;
            int outPut = 0;
            if (dept != "PAE" && role != "Edit")
            {
                return Json(new { status = false, msg = "You do not have permition!" }, JsonRequestBehavior.AllowGet);
            }

            if (ids.Length > 0)
            {
                foreach (var id in ids)
                {

                    Die1 exitDie = db.Die1.Find(int.Parse(id));

                    if (Session["Admin"].ToString() == "Admin")
                    {

                        //1. Xoa commonDie1 table
                        var coms = db.CommonDie1.Where(x => x.DieID == exitDie.DieID).ToList();
                        foreach (var item in coms)
                        {
                            item.Active = false;
                            db.Entry(item).State = EntityState.Modified;
                            db.SaveChanges();
                        }
                        exitDie.isCancel = true;
                        exitDie.isOfficial = false;
                        exitDie.Active = false;
                        exitDie.Genaral_Information = DateTime.Now.ToString("MM/dd/yyyy ") + Session["Name"].ToString() + ": Cancel this die " + System.Environment.NewLine + exitDie.Genaral_Information;
                        //exitDie = commonFunction.updateDieStatus(exitDie);
                        db.Entry(exitDie).State = EntityState.Modified;
                        db.SaveChanges();
                        outPut = outPut + 1;
                        //exitDie = commonFunction.updateDieStatus(exitDie);

                    }
                    if (exitDie.DieClassify != "MP")
                    {
                        exitDie.isCancel = true;
                        exitDie.isOfficial = false;
                        exitDie.Genaral_Information = DateTime.Now.ToString("MM/dd/yyyy ") + Session["Name"].ToString() + ": Cancel this die " + System.Environment.NewLine + exitDie.Genaral_Information;
                        //exitDie = commonFunction.updateDieStatus(exitDie);
                        db.Entry(exitDie).State = EntityState.Modified;
                        db.SaveChanges();
                        outPut = outPut + 1;
                    }

                }
            }
            return Json(new { inPut = inPut, success = outPut }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult SaveData(Die1 record, string FileName, HttpPostedFileBase file, HttpPostedFileBase FAAttachment, HttpPostedFileBase ToAttachment)
        {
            //db.Configuration.ProxyCreationEnabled = false;
            //var findRecord = db.Die_Launch_Management.Find(record.RecordID);
            Die1 findRecord = db.Die1.Find(record.DieID);
            // Check Quyền Update
            var dept = Session["Dept"].ToString();
            var role = Session["Die_Lauch_Role"].ToString();
            var admin = Session["Role"].ToString() == "admin";
            var status = false;
            if (role != "Edit")
            {
                var data1 = new
                {
                    status = false,
                    record = findRecord
                };
                return Json(data1, JsonRequestBehavior.AllowGet);
            }

            var today = DateTime.Now;

            //*********************************************
            // Giữ lại pending_status trước khi được update
            var currentPendingStatus = commonFunction.getPendingAndDeptResponse(findRecord).Pending_Status;
            //**********************************************

            // update nếu có value chuyển lên.
            findRecord.Step = record.Step == null ? findRecord.Step : record.Step;

            findRecord.Step = record.Step == null ? findRecord.Step : record.Step;
            findRecord.Rank = record.Rank == null ? findRecord.Rank : record.Rank;
            // Kiển trả part No hợp lệ hay ko?
            if (!String.IsNullOrEmpty(record.PartNoOriginal) && !String.IsNullOrEmpty(record.Die_Code) && findRecord.DieClassify != "MP")
            {
                var dieCode = record.Die_Code.ToUpper().Trim();
                var partNo = record.PartNoOriginal.ToUpper().Trim();
                if (partNo.Length == 12 && dieCode.Length == 3) //RC5-1234-000
                {

                    var dieNo = partNo + "-" + dieCode + "-001"; //RC5-1234-000-11A-001
                    var existDie = db.Die1.Where(x => x.DieID != findRecord.DieID && x.DieNo == dieNo && x.Active != false).Count();
                    if (existDie == 0) // chưa tồn tại
                    {
                        findRecord.PartNoOriginal = partNo;
                        findRecord.Die_Code = dieCode;
                        findRecord.DieNo = dieNo;
                    }
                    findRecord.DieClassify = commonFunction.isRenewOrAddOrMT(dieCode);
                }

            }




            findRecord.ModelID = record.ModelID > 0 ? record.ModelID : findRecord.ModelID;
            findRecord.SupplierID = record.SupplierID >= 0 ? record.SupplierID : findRecord.SupplierID;
            findRecord.Texture = record.Texture == null ? findRecord.Texture : record.Texture;

            // Genaral được update nối đôi
            if (!String.IsNullOrWhiteSpace(record.Genaral_Information))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + " :" + record.Genaral_Information + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.Decision_Date = record.Decision_Date == null ? findRecord.Decision_Date : record.Decision_Date;
            if (String.IsNullOrWhiteSpace(findRecord.Select_Supplier_Date) && !String.IsNullOrWhiteSpace(record.Select_Supplier_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": Select Supplier on " + record.Select_Supplier_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.Select_Supplier_Date = record.Select_Supplier_Date == null ? findRecord.Select_Supplier_Date : record.Select_Supplier_Date;

            if (String.IsNullOrWhiteSpace(findRecord.QTN_Sub_Date) && !String.IsNullOrWhiteSpace(record.QTN_Sub_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": QTN submit on " + record.QTN_Sub_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.QTN_Sub_Date = record.QTN_Sub_Date == null ? findRecord.QTN_Sub_Date : record.QTN_Sub_Date;

            if (String.IsNullOrWhiteSpace(findRecord.QTN_App_Date) && !String.IsNullOrWhiteSpace(record.QTN_App_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": QTN App on " + record.QTN_App_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.QTN_App_Date = record.QTN_App_Date == null ? findRecord.QTN_App_Date : record.QTN_App_Date;

            findRecord.Need_Use_Date = record.Need_Use_Date == null ? findRecord.Need_Use_Date : record.Need_Use_Date;
            findRecord.Target_OK_Date = record.Target_OK_Date == null ? findRecord.Target_OK_Date : record.Target_OK_Date;
            findRecord.Inv_Idea = record.Inv_Idea == null ? findRecord.Inv_Idea : record.Inv_Idea;
            findRecord.Inv_FB_To = record.Inv_FB_To == null ? findRecord.Inv_FB_To : record.Inv_FB_To;
            findRecord.Inv_Result = record.Inv_Result == null ? findRecord.Inv_Result : record.Inv_Result;
            findRecord.Inv_Cost_Down = record.Inv_Cost_Down == null ? findRecord.Inv_Cost_Down : record.Inv_Cost_Down;

            if (String.IsNullOrWhiteSpace(findRecord.DFM_Sub_Date) && !String.IsNullOrWhiteSpace(record.DFM_Sub_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM submit on " + record.DFM_Sub_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.DFM_Sub_Date = record.DFM_Sub_Date == null ? findRecord.DFM_Sub_Date : record.DFM_Sub_Date;

            if (String.IsNullOrWhiteSpace(findRecord.DFM_PAE_Check_Date) && !String.IsNullOrWhiteSpace(record.DFM_PAE_Check_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM submit on " + record.DFM_PAE_Check_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.DFM_PAE_Check_Date = record.DFM_PAE_Check_Date == null ? findRecord.DFM_PAE_Check_Date : record.DFM_PAE_Check_Date;

            if (String.IsNullOrWhiteSpace(findRecord.DFM_PE_Check_Date) && !String.IsNullOrWhiteSpace(record.DFM_PE_Check_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM submit on " + record.DFM_PE_Check_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.DFM_PE_Check_Date = record.DFM_PE_Check_Date == null ? findRecord.DFM_PE_Check_Date : record.DFM_PE_Check_Date;

            if (String.IsNullOrWhiteSpace(findRecord.DFM_PE_App_Date) && !String.IsNullOrWhiteSpace(record.DFM_PE_App_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM submit on " + record.DFM_PE_App_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.DFM_PE_App_Date = record.DFM_PE_App_Date == null ? findRecord.DFM_PE_App_Date : record.DFM_PE_App_Date;

            if (String.IsNullOrWhiteSpace(findRecord.DFM_PAE_App_Date) && !String.IsNullOrWhiteSpace(record.DFM_PAE_App_Date))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM submit on " + record.DFM_PAE_App_Date + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.DFM_PAE_App_Date = record.DFM_PAE_App_Date == null ? findRecord.DFM_PAE_App_Date : record.DFM_PAE_App_Date;

            findRecord.CoreCavMaterial = record.CoreCavMaterial == null ? findRecord.CoreCavMaterial : record.CoreCavMaterial;
            findRecord.SliderMaterial = record.SliderMaterial == null ? findRecord.SliderMaterial : record.SliderMaterial;
            findRecord.LifterMaterial = record.LifterMaterial == null ? findRecord.LifterMaterial : record.LifterMaterial;


            findRecord.PunchBackingPlate = record.PunchBackingPlate == null ? findRecord.PunchBackingPlate : record.PunchBackingPlate;
            findRecord.PunchPlate = record.PunchPlate == null ? findRecord.PunchPlate : record.PunchPlate;
            findRecord.StripperBackingPlate = record.StripperBackingPlate == null ? findRecord.StripperBackingPlate : record.StripperBackingPlate;
            findRecord.StripperPlate = record.StripperPlate == null ? findRecord.StripperPlate : record.StripperPlate;
            findRecord.DiePlate = record.DiePlate == null ? findRecord.DiePlate : record.DiePlate;
            findRecord.DieBackingPlate = record.DieBackingPlate == null ? findRecord.DieBackingPlate : record.DieBackingPlate;
            findRecord.Punch = record.Punch == null ? findRecord.Punch : record.Punch;
            findRecord.InsertBlock = record.InsertBlock == null ? findRecord.InsertBlock : record.InsertBlock;
            findRecord.isStacking = record.isStacking == null ? findRecord.isStacking : record.isStacking;
            findRecord.isShaftKashime = record.isShaftKashime == null ? findRecord.isShaftKashime : record.isShaftKashime;
            findRecord.isBurringKashime = record.isBurringKashime == null ? findRecord.isBurringKashime : record.isBurringKashime;
            findRecord.PXNoOfComponent = record.PXNoOfComponent == null ? findRecord.PXNoOfComponent : record.PXNoOfComponent;




            findRecord.HotRunner = record.HotRunner == null ? findRecord.HotRunner : record.HotRunner;
            findRecord.GateType = record.GateType == null ? findRecord.GateType : record.GateType;
            findRecord.MCsize = record.MCsize == null ? findRecord.MCsize : record.MCsize;
            findRecord.CavQuantity = record.CavQuantity == null ? findRecord.CavQuantity : record.CavQuantity;
            findRecord.DieMakeLocation = record.DieMakeLocation == null ? findRecord.DieMakeLocation : record.DieMakeLocation;
            findRecord.DieMaker = record.DieMaker == null ? findRecord.DieMaker : record.DieMaker;
            findRecord.Family_Die_With = record.Family_Die_With == null ? findRecord.Family_Die_With : record.Family_Die_With;
            findRecord.Common_Part_With = record.Common_Part_With == null ? findRecord.Common_Part_With : record.Common_Part_With;
            findRecord.SpecialSpec = record.SpecialSpec == null ? findRecord.SpecialSpec : record.SpecialSpec;
            findRecord.DSUM_Idea = record.DSUM_Idea == null ? findRecord.DSUM_Idea : record.DSUM_Idea;
            // Tạm thòi giữ lại
            // sẽ được tự đông update PO issue và PO App
            findRecord.PO_Issue_Date = record.PO_Issue_Date == null ? findRecord.PO_Issue_Date : record.PO_Issue_Date;
            findRecord.PODate = record.PODate == null ? findRecord.PODate : record.PODate;
            //**************************

            findRecord.Design_Check_Plan = record.Design_Check_Plan == null ? findRecord.Design_Check_Plan : record.Design_Check_Plan;
            findRecord.Design_Check_Actual = record.Design_Check_Actual == null ? findRecord.Design_Check_Actual : record.Design_Check_Actual;
            findRecord.Design_Check_Result = record.Design_Check_Result == null ? findRecord.Design_Check_Result : record.Design_Check_Result;
            findRecord.NoOfPoit_Not_FL_DMF = record.NoOfPoit_Not_FL_DMF == null ? findRecord.NoOfPoit_Not_FL_DMF : record.NoOfPoit_Not_FL_DMF;
            findRecord.NoOfPoint_Not_FL_Spec = record.NoOfPoint_Not_FL_Spec == null ? findRecord.NoOfPoint_Not_FL_Spec : record.NoOfPoint_Not_FL_Spec;
            findRecord.NoOfPoint_Kaizen = record.NoOfPoint_Kaizen == null ? findRecord.NoOfPoint_Kaizen : record.NoOfPoint_Kaizen;
            findRecord.JIG_Using = record.JIG_Using == null ? findRecord.JIG_Using : record.JIG_Using;
            findRecord.JIG_No = record.JIG_No == null ? findRecord.JIG_No : record.JIG_No;
            findRecord.JIG_Check_Plan = record.JIG_Check_Plan == null ? findRecord.JIG_Check_Plan : record.JIG_Check_Plan;
            findRecord.JIG_Check_Result = record.JIG_Check_Result == null ? findRecord.JIG_Check_Result : record.JIG_Check_Result;
            findRecord.JIG_ETA_Supplier = record.JIG_ETA_Supplier == null ? findRecord.JIG_ETA_Supplier : record.JIG_ETA_Supplier;
            findRecord.JIG_Status = record.JIG_Status == null ? findRecord.JIG_Status : record.JIG_Status;
            findRecord.T0_Plan = record.T0_Plan == null ? findRecord.T0_Plan : record.T0_Plan;
            findRecord.T0_Actual = record.T0_Actual == null ? findRecord.T0_Actual : record.T0_Actual;

            if (String.IsNullOrWhiteSpace(findRecord.T0_Result) && !String.IsNullOrWhiteSpace(record.T0_Result))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": TO trial result " + record.T0_Result + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.T0_Result = record.T0_Result == null ? findRecord.T0_Result : record.T0_Result;

            findRecord.T0_Solve_Method = record.T0_Solve_Method == null ? findRecord.T0_Solve_Method : record.T0_Solve_Method;
            findRecord.T0_Solve_Result = record.T0_Solve_Result == null ? findRecord.T0_Solve_Result : record.T0_Solve_Result;
            findRecord.Texture_Meeting_Date = record.Texture_Meeting_Date == null ? findRecord.Texture_Meeting_Date : record.Texture_Meeting_Date;
            findRecord.Texture_Go_Date = record.Texture_Go_Date == null ? findRecord.Texture_Go_Date : record.Texture_Go_Date;
            findRecord.S0_Plan = record.S0_Plan == null ? findRecord.S0_Plan : record.S0_Plan;
            findRecord.S0_Result = record.S0_Result == null ? findRecord.S0_Result : record.S0_Result;
            findRecord.S0_Solve_Method = record.S0_Solve_Method == null ? findRecord.S0_Solve_Method : record.S0_Solve_Method;
            findRecord.S0_solve_Result = record.S0_solve_Result == null ? findRecord.S0_solve_Result : record.S0_solve_Result;
            findRecord.Texture_App_Date = record.Texture_App_Date == null ? findRecord.Texture_App_Date : record.Texture_App_Date;
            findRecord.Texture_Internal_App_Result = record.Texture_Internal_App_Result == null ? findRecord.Texture_Internal_App_Result : record.Texture_Internal_App_Result;
            findRecord.Texture_JP_HP_App_Result = record.Texture_JP_HP_App_Result == null ? findRecord.Texture_JP_HP_App_Result : record.Texture_JP_HP_App_Result;
            findRecord.Texture_Note = record.Texture_Note == null ? findRecord.Texture_Note : record.Texture_Note;
            findRecord.PreKK_Plan = record.PreKK_Plan == null ? findRecord.PreKK_Plan : record.PreKK_Plan;
            findRecord.PreKK_Actual = record.PreKK_Actual == null ? findRecord.PreKK_Actual : record.PreKK_Actual;
            findRecord.PreKK_Result = record.PreKK_Result == null ? findRecord.PreKK_Result : record.PreKK_Result;



            findRecord.FA_Plan = record.FA_Plan == null ? findRecord.FA_Plan : record.FA_Plan;
            if (!String.IsNullOrWhiteSpace(record.FA_Result))
            {
                if (findRecord.FA_Result != record.FA_Result)
                {
                    findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": FA Result/Position " + record.FA_Result + System.Environment.NewLine + findRecord.Genaral_Information;
                    if (record.FA_Result_Date == null)
                    {
                        record.FA_Result_Date = today;
                    }

                }
            }

            findRecord.FA_Result = record.FA_Result == null ? findRecord.FA_Result : record.FA_Result;
            findRecord.FA_Result_Date = record.FA_Result_Date == null ? findRecord.FA_Result_Date : record.FA_Result_Date;
            findRecord.FA_Problem = record.FA_Problem == null ? findRecord.FA_Problem : record.FA_Problem;

            if (String.IsNullOrWhiteSpace(findRecord.FA_Action_Improve) && !String.IsNullOrWhiteSpace(record.FA_Action_Improve))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": Action Improve FA " + record.FA_Action_Improve + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            findRecord.FA_Action_Improve = record.FA_Action_Improve == null ? findRecord.FA_Action_Improve : record.FA_Action_Improve;
            findRecord.MT1_Date = record.MT1_Date == null ? findRecord.MT1_Date : record.MT1_Date;
            findRecord.MT1_Gather_Date = record.MT1_Gather_Date == null ? findRecord.MT1_Gather_Date : record.MT1_Gather_Date;
            findRecord.MT1_Problem = record.MT1_Problem == null ? findRecord.MT1_Problem : record.MT1_Problem;
            findRecord.MT1_Remark = record.MT1_Remark == null ? findRecord.MT1_Remark : record.MT1_Remark;
            findRecord.MTF_Date = record.MTF_Date == null ? findRecord.MTF_Date : record.MTF_Date;
            findRecord.MTF_Gather_Date = record.MTF_Gather_Date == null ? findRecord.MTF_Gather_Date : record.MTF_Gather_Date;
            findRecord.MTF_Problem = record.MTF_Problem == null ? findRecord.MTF_Problem : record.MTF_Problem;
            findRecord.MTF_Remark = record.MTF_Remark == null ? findRecord.MTF_Remark : record.MTF_Remark;

            if (String.IsNullOrWhiteSpace(findRecord.TVP_No) && !String.IsNullOrWhiteSpace(record.TVP_No))
            {
                findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": Issue TVP " + record.TVP_No + System.Environment.NewLine + findRecord.Genaral_Information;
            }
            //findRecord.TVP_No = record.TVP_No == null ? findRecord.TVP_No : record.TVP_No;
            //findRecord.ERI_No = record.ERI_No == null ? findRecord.ERI_No : record.ERI_No;
            //if (!String.IsNullOrWhiteSpace(record.TVP_Result))
            //{
            //    if (findRecord.TVP_Result != record.TVP_Result)
            //    {
            //        findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": TVP Result/Position " + record.TVP_Result + System.Environment.NewLine + findRecord.Genaral_Information;
            //    }
            //}

            //findRecord.TVP_Result = record.TVP_Result == null ? findRecord.TVP_Result : record.TVP_Result;
            //findRecord.TVP_Result_Date = record.TVP_Result_Date == null ? findRecord.TVP_Result_Date : record.TVP_Result_Date;
            //findRecord.TVP_Remark = record.TVP_Remark == null ? findRecord.TVP_Remark : record.TVP_Remark;
            //findRecord.PCAR_Date = record.PCAR_Date == null ? findRecord.PCAR_Date : record.PCAR_Date;
            //if (!String.IsNullOrWhiteSpace(record.PCAR_Result))
            //{
            //    if (findRecord.PCAR_Result != record.PCAR_Result)
            //    {
            //        findRecord.Genaral_Information = today.ToString("MM/dd/yyyy") + ": TVP Result/Position " + record.PCAR_Result + System.Environment.NewLine + findRecord.Genaral_Information;
            //    }
            //}

            //findRecord.PCAR_Result = record.PCAR_Result == null ? findRecord.PCAR_Result : record.PCAR_Result;
            findRecord.First_Lot_Date = record.First_Lot_Date == null ? findRecord.First_Lot_Date : record.First_Lot_Date;

            // Sửa lại code khi app
            findRecord.EditBy = Session["Name"].ToString();
            findRecord.EditDate = today;




            // Làm gì khi pending status thay đôi

            // 1. Tính Số lần submit
            //******Làm thế nào để tính số lần submit Time
            // Mặc định là 1 lần
            // Nếu nhảy sang trạng thái W-FA-Resubmit => + 1 => nhưng chỉ hiện thị Số lần submitFA - 1

            // 2. Record lại thời gian thay đổi trạng thái


            {
                //*********************************************
                // pending_status sau khi được update
                var newPendingStatus = commonFunction.getPendingAndDeptResponse(findRecord);
                //**********************************************
                if (findRecord.FA_Sub_Time == null || findRecord.FA_Sub_Time == 0)
                {
                    findRecord.FA_Sub_Time = 1;
                }
                if (currentPendingStatus != newPendingStatus.Pending_Status)
                {
                    //1. Tính số lần submit FA 
                    if (newPendingStatus.Pending_Status.Contains("FA_ReSubmit"))
                    {
                        findRecord.FA_Sub_Time = findRecord.FA_Sub_Time + 1;
                    }
                    findRecord.Latest_Pending_Status_Changed = today;
                }
                // Luu status
                // findRecord = commonFunction.updateDieStatus(findRecord);

            }


            // Phân luu Att file
            if (ToAttachment != null)
            {
                var fileExt = Path.GetExtension(ToAttachment.FileName);
                // Code luu file vào folder
                var fileName = "Attachment_T0Trial_" + findRecord.DieNo + today.ToString("yyyy-MM-dd-HHmmss") + fileExt;
                var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
                ToAttachment.SaveAs(path);
                Models.Attachment newAtt = new Models.Attachment
                {
                    DieID = findRecord.DieID,
                    FileName = fileName,
                    Clasify = "T0",
                    CreateBy = Session["Name"].ToString(),
                    CreateDate = today
                };
                db.Attachments.Add(newAtt);
                db.SaveChanges();
            }
            if (FAAttachment != null)
            {
                var fileExt = Path.GetExtension(FAAttachment.FileName);
                // Code luu file vào folder
                var fileName = "Attachment_FA_Improve_" + findRecord.DieNo + today.ToString("yyyy-MM-dd-HHmmss") + fileExt;
                var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
                FAAttachment.SaveAs(path);
                Models.Attachment newAtt = new Models.Attachment
                {
                    DieID = findRecord.DieID,
                    FileName = fileName,
                    Clasify = "FA Improve",
                    CreateBy = Session["Name"].ToString(),
                    CreateDate = today
                };
                db.Attachments.Add(newAtt);
                db.SaveChanges();
            }

            if (file != null && FileName != null)
            {
                var fileExt = Path.GetExtension(file.FileName);
                // Code luu file vào folder
                var fileName = FileName + "_" + findRecord.DieNo + today.ToString("yyyy-MM-dd-HHmmss") + fileExt;
                var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
                file.SaveAs(path);
                Models.Attachment newAtt = new Models.Attachment
                {
                    DieID = findRecord.DieID,
                    FileName = fileName,
                    Clasify = "Others",
                    CreateBy = Session["Name"].ToString(),
                    CreateDate = today
                };
                db.Attachments.Add(newAtt);
                db.SaveChanges();
            }





            db.Entry(findRecord).State = EntityState.Modified;
            db.SaveChanges();
            // send Email


            List<CommonDie1> result = new List<CommonDie1>();
            var cm = db.CommonDie1.Where(x => x.DieID == findRecord.DieID && x.Active != false).FirstOrDefault();
            result.Add(cm);
            status = true;
            var data = new
            {
                status = status,
                record = dataReturnView(result)
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EditAttachment(int AttachID, HttpPostedFileBase file)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Die_Lauch_Role"].ToString();
            if (role != "Edit")
            {
                return Json(new { status = false, msg = "You do not have permition!" }, JsonRequestBehavior.AllowGet);
            }

            var status = false;
            var att = db.Attachments.Find(AttachID);
            var today = DateTime.Now;
            if (file != null)
            {
                var fileExt = Path.GetExtension(file.FileName);
                // Code luu file vào folder
                var oldName = Path.GetFileNameWithoutExtension(att.FileName);
                oldName = oldName.Remove(oldName.Length - 17);
                var fileName = oldName + today.ToString("yyyy-MM-dd-HHmmss") + "(revised)" + fileExt;
                var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
                file.SaveAs(path);

                att.FileName = fileName;
                att.CreateDate = today;
                att.CreateBy = Session["Name"].ToString();
                db.Entry(att).State = EntityState.Modified;
                db.SaveChanges();
                status = true;
            }

            var output = new
            {
                status = status,
                DieID = att.DieID
            };

            return Json(output, JsonRequestBehavior.AllowGet);
        }


        public async Task<JsonResult> dataForChart()
        {
            string[] cfDeptRes = { "PAE", "PE1", "PUR", "PUS", "Supplier" };
            string[] cfCategory = { "MT", "Renewal", "Additional" };
            DieLaunchingWarningConfig[] cfPending = db.DieLaunchingWarningConfigs.ToArray();

            List<Supplier> suppliers = new List<Supplier>();
            List<object> dataStatus = new List<object>();
            List<object> dataSupplier = new List<object>();
            List<string> lablesStatus = new List<string>();
            List<string> lablesSuppliers = new List<string>();
            List<object> dataRes = new List<object>();

            var task = await Task.Factory.StartNew(() =>
            {
                var allDie = db.Die1.Where(x => x.Active != false && !x.DieClassify.Contains("MP")).ToList();
                foreach (var sup in db.Suppliers.ToList())
                {
                    Die1 d = allDie.Where(x => x.SupplierID == sup.SupplierID).FirstOrDefault();
                    if (d != null)
                    {
                        suppliers.Add(sup);
                    }
                }
                for (int i = 0; i < cfCategory.Length; i++)
                {
                    List<int> output = new List<int>();
                    for (int j = 0; j < cfDeptRes.Length; j++)
                    {
                        string dept = cfDeptRes[j];
                        string cate = cfCategory[i];
                        int c = allDie.Where(x => x.DieClassify.Contains(cate) && commonFunction.getPendingAndDeptResponse(x).Dept_Respone.Contains(dept)).Count();
                        output.Add(c);

                    }

                    object n = new
                    {
                        name = cfCategory[i],
                        data = output
                    };
                    dataRes.Add(n);


                    // Dem so luong them tung trang thais

                    List<int> outPutdataStatus = new List<int>();
                    lablesStatus.Clear();
                    for (int c = 0; c < cfPending.Count(); c++)
                    {

                        string status = cfPending[c].Status;
                        string cate = cfCategory[i];
                        int dem = allDie.Where(x => x.DieClassify.Contains(cate) && commonFunction.getPendingAndDeptResponse(x).Pending_Status == status).Count();
                        outPutdataStatus.Add(dem);
                        lablesStatus.Add(status);
                    }
                    dataStatus.Add(new
                    {
                        name = cfCategory[i],
                        data = outPutdataStatus
                    });


                    // Dep theo maker


                    List<int> outPutdataSupplier = new List<int>();
                    lablesSuppliers.Clear();
                    for (int d = 0; d < suppliers.Count(); d++)
                    {

                        Supplier supplier = suppliers[d];
                        string cate = cfCategory[i];
                        int dem = allDie.Where(x => x.DieClassify.Contains(cate) && x.SupplierID == supplier.SupplierID && x.isCancel != true).Count();

                        outPutdataSupplier.Add(dem);
                        lablesSuppliers.Add(supplier.SupplierCode);


                    }

                    dataSupplier.Add(new
                    {
                        name = cfCategory[i],
                        data = outPutdataSupplier
                    });
                }


                object result = new
                {
                    DeptRespone = new
                    {
                        data = dataRes,
                        lables = cfDeptRes
                    },

                    MakingStatus = new
                    {
                        data = dataStatus,
                        lables = lablesStatus
                    },
                    SupplierCapacity = new
                    {
                        data = dataSupplier,
                        lables = lablesSuppliers
                    }
                };
                return Json(result, JsonRequestBehavior.AllowGet);
            });

            return Json(task, JsonRequestBehavior.AllowGet);
        }

        public IEnumerable pendingStatusCategory()
        {
            IEnumerable Pending_list = new[]
           {
                new { value = "1.W-SelectSupplier" ,show = "1.W-SelectSupplier"},
                new { value = "2.W-QTN_Sub" ,show = "2.W-QTN_Sub"},
                new { value = "3.W-QTN_App." ,show = "3.W-QTN_App."},
                new { value = "4.W-DFM_Sub" ,show = "4.W-DFM_Sub"},
                new { value = "5.W-DFM_PAE_Check" ,show = "5.W-DFM_PAE_Check"},
                new { value = "6.W-DFM_PE1_Check" ,show = "6.W-DFM_PE1_Check"},
                new { value = "7.W-DFM_PE1_App" ,show = "7.W-DFM_PE1_App"},
                new { value = "8.W-DFM_PAE_App" ,show = "8.W-DFM_PAE_App"},
                new { value = "9.W-MR_Issue" ,show = "9.W-MR_Issue"},
                new { value = "10.W-MR_App" ,show = "10.W-MR_App"},
                new { value = "11.W-PO_Issue" ,show = "11.W-PO_Issue"},
                new { value = "12.W-PO_App" ,show = "12.W-PO_App"},
                new { value = "13.W-T0_Plan" ,show = "13.W-T0_Plan"},
                new { value = "14.W-T0_trial" ,show = "14.W-T0_trial"},
                new { value = "15.PlzConfirmT0Result" ,show = "15.PlzConfirmT0Result"},
                new { value = "16.W-FA_Plan" ,show = "16.W-FA_Plan"},
                new { value = "17.W-FA_Submit" ,show = "17.W-FA_Submit"},
                new { value = "18.PlzConfirmFASubmit?" ,show = "18.PlzConfirmFASubmit?"},
                new { value = "19.W-FA_Result" ,show = "19.W-FA_Result"},
                new { value = "20.W-RepairMethod" ,show = "20.W-RepairMethod"},
                new { value = "21.W-FA_ReSubmit" ,show = "21.W-FA_ReSubmit"},
                new { value = "22.W-TVP_Issue" ,show = "22.W-TVP_Issue"},
                new { value = "23.W-TVP_Result" ,show = "23.W-TVP_Result"},
                new { value = "24.W-ReTVP_Result" ,show = "24.W-ReTVP_Result"},
                new { value = "25.W-PCAR_Result" ,show = "25.W-PCAR_Result"},
                new { value = "26.Done" ,show = "26.Done"},
            };
            return Pending_list;
        }


        public IEnumerable statusCategory()
        {
            IEnumerable status = new[]
           {
                new { value = "Late" ,show = "Late"},
                new { value = "Earlier" ,show = "Earlier"},
                new { value = "OnTime" ,show = "OnTime"},
                new { value = "OnProgres" ,show = "OnProgres"},
                new { value = "Pending" ,show = "Pending"},
                new { value = "Warning" ,show = "Warning"},
                new { value = "Plz input Target OK" ,show = "Plz input Target OK"},
                new { value = "Plz Input FA Result Date" ,show = "Plz Input FA Result Date"},
            };
            return status;
        }


        public IEnumerable deptResponse()
        {
            IEnumerable status = new[]
           {
                new { value = "PUR" ,show = "PUR"},
                new { value = "PUS" ,show = "PUS"},
                new { value = "MQA" ,show = "MQA"},
                new { value = "PQA" ,show = "PQA"},
                new { value = "PE1" ,show = "PE1"},
                new { value = "PAE" ,show = "PAE"},
                new { value = "PDC" ,show = "PDC"},
                new { value = "OQA" ,show = "OQA"},
            };
            return status;
        }

        public IEnumerable clasify_list()
        {
            IEnumerable clasify_list = new[]
           {
                 new { value = "MP", show = "MP" },
                 new { value = "MT", show = "MT" },
                  new { value = "Renewal", show = "Renewal" },
                  new { value = "Additional", show = "Additional" },

            };
            return clasify_list;
        }





        //public Object exportToControlList(List<CommonDie1> records)
        //{
        //    string handle = Guid.NewGuid().ToString();
        //    MemoryStream output = new MemoryStream();
        //    using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/UpdateDieInfo/Format_Die_Master_List01.xlsx"))))
        //    {
        //        ExcelWorksheet sheet = package.Workbook.Worksheets.First();
        //        int rowId = 7;
        //        int i = 1;

        //        foreach (var x in records)
        //        {
        //            var pend = commonFunction.getPendingAndDeptResponse(x.Die1);
        //            sheet.Cells["A" + rowId.ToString()].Value = i;
        //            sheet.Cells["B" + rowId.ToString()].Value = x.Die1.DieClassify;
        //            sheet.Cells["C" + rowId.ToString()].Value = x.Die1.Step;
        //            sheet.Cells["D" + rowId.ToString()].Value = x.Die1.Rank;
        //            sheet.Cells["E" + rowId.ToString()].Value = x.PartNo;
        //            sheet.Cells["F" + rowId.ToString()].Value = x.Parts1.PartName;
        //            sheet.Cells["G" + rowId.ToString()].Value = x.Die1.ProcessCodeCalogory.Type;
        //            sheet.Cells["H" + rowId.ToString()].Value = x.Die1.Die_Code;
        //            sheet.Cells["I" + rowId.ToString()].Value = x.Die1.DieNo;
        //            sheet.Cells["J" + rowId.ToString()].Value = x.Die1.ModelID > 0 ? x.Die1.ModelList.ModelName : "";
        //            sheet.Cells["K" + rowId.ToString()].Value = x.Die1.SupplierID >= 0 ? x.Die1.Supplier.SupplierCode : "";
        //            sheet.Cells["L" + rowId.ToString()].Value = x.Die1.SupplierID >= 0 ? x.Die1.Supplier.SupplierName : "";
        //            sheet.Cells["M" + rowId.ToString()].Value = pend.Progress;
        //            sheet.Cells["N" + rowId.ToString()].Value = commonFunction.getStatus(x.Die1);
        //            sheet.Cells["O" + rowId.ToString()].Value = pend.Pending_Status;
        //            sheet.Cells["P" + rowId.ToString()].Value = pend.Dept_Respone;
        //            sheet.Cells["Q" + rowId.ToString()].Value = commonFunction.getWarning(x.Die1);
        //            sheet.Cells["R" + rowId.ToString()].Value = x.Die1.Genaral_Information;
        //            sheet.Cells["S" + rowId.ToString()].Value = x.Die1.Decision_Date;
        //            sheet.Cells["T" + rowId.ToString()].Value = x.Die1.Select_Supplier_Date;
        //            sheet.Cells["U" + rowId.ToString()].Value = x.Die1.QTN_Sub_Date;
        //            sheet.Cells["V" + rowId.ToString()].Value = x.Die1.QTN_App_Date;
        //            sheet.Cells["W" + rowId.ToString()].Value = x.Die1.Need_Use_Date;
        //            sheet.Cells["X" + rowId.ToString()].Value = x.Die1.Target_OK_Date;
        //            sheet.Cells["Y" + rowId.ToString()].Value = x.Die1.Inv_Idea;
        //            sheet.Cells["Z" + rowId.ToString()].Value = x.Die1.Inv_FB_To;
        //            sheet.Cells["AA" + rowId.ToString()].Value = x.Die1.Inv_Result;
        //            sheet.Cells["AB" + rowId.ToString()].Value = x.Die1.Inv_Cost_Down;
        //            sheet.Cells["AC" + rowId.ToString()].Value = x.Die1.DFM_Sub_Date;
        //            sheet.Cells["AD" + rowId.ToString()].Value = x.Die1.DFM_PAE_Check_Date;
        //            sheet.Cells["AE" + rowId.ToString()].Value = x.Die1.DFMPE1Checked;
        //            sheet.Cells["AF" + rowId.ToString()].Value = x.Die1.DFM_PE_App_Date;
        //            sheet.Cells["AG" + rowId.ToString()].Value = x.Die1.DFM_PAE_App_Date;
        //            sheet.Cells["AH" + rowId.ToString()].Value = x.Die1.CoreCavMaterial;
        //            sheet.Cells["AI" + rowId.ToString()].Value = x.Die1.SliderMaterial;
        //            sheet.Cells["AJ" + rowId.ToString()].Value = x.Die1.LifterMaterial;
        //            sheet.Cells["AK" + rowId.ToString()].Value = x.Die1.Texture == true ? "Yes" : x.Die1.Texture == false ? "No" : "";
        //            sheet.Cells["AL" + rowId.ToString()].Value = x.Die1.HotRunner;
        //            sheet.Cells["AM" + rowId.ToString()].Value = x.Die1.GateType;
        //            sheet.Cells["AN" + rowId.ToString()].Value = x.Die1.MCsize;
        //            sheet.Cells["AO" + rowId.ToString()].Value = x.Die1.CavQuantity;
        //            sheet.Cells["AP" + rowId.ToString()].Value = x.Die1.DieMakeLocation;
        //            sheet.Cells["AQ" + rowId.ToString()].Value = x.Die1.DieMaker;
        //            sheet.Cells["AR" + rowId.ToString()].Value = x.Die1.Family_Die_With;
        //            sheet.Cells["AS" + rowId.ToString()].Value = x.Die1.Common_Part_With;
        //            sheet.Cells["AT" + rowId.ToString()].Value = x.Die1.SpecialSpec;
        //            sheet.Cells["AU" + rowId.ToString()].Value = x.Die1.DSUM_Idea;
        //            sheet.Cells["AV" + rowId.ToString()].Value = x.Die1.MR_Request_Date.HasValue ? x.Die1.MR_Request_Date.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["AW" + rowId.ToString()].Value = x.Die1.MR_App_Date.HasValue ? x.Die1.MR_App_Date.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["AX" + rowId.ToString()].Value = x.Die1.PO_Issue_Date.HasValue ? x.Die1.PO_Issue_Date.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["AY" + rowId.ToString()].Value = x.Die1.PODate.HasValue ? x.Die1.PODate.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["AZ" + rowId.ToString()].Value = x.Die1.Design_Check_Plan;
        //            sheet.Cells["BA" + rowId.ToString()].Value = x.Die1.Design_Check_Actual;
        //            sheet.Cells["BB" + rowId.ToString()].Value = x.Die1.Design_Check_Result;
        //            sheet.Cells["BC" + rowId.ToString()].Value = x.Die1.NoOfPoit_Not_FL_DMF;
        //            sheet.Cells["BD" + rowId.ToString()].Value = x.Die1.NoOfPoint_Not_FL_Spec;
        //            sheet.Cells["BE" + rowId.ToString()].Value = x.Die1.NoOfPoint_Kaizen;
        //            sheet.Cells["BF" + rowId.ToString()].Value = x.Die1.JIG_Using == true ? "Yes" : x.Die1.Texture == false ? "No" : ""; ;
        //            sheet.Cells["BG" + rowId.ToString()].Value = x.Die1.JIG_No;
        //            sheet.Cells["BH" + rowId.ToString()].Value = x.Die1.JIG_Check_Plan;
        //            sheet.Cells["BI" + rowId.ToString()].Value = x.Die1.JIG_Check_Result;
        //            sheet.Cells["BJ" + rowId.ToString()].Value = x.Die1.JIG_ETA_Supplier;
        //            sheet.Cells["BK" + rowId.ToString()].Value = x.Die1.JIG_Status;
        //            sheet.Cells["BL" + rowId.ToString()].Value = x.Die1.T0_Plan.HasValue ? x.Die1.T0_Plan.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["BM" + rowId.ToString()].Value = x.Die1.T0_Actual.HasValue ? x.Die1.T0_Actual.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["BN" + rowId.ToString()].Value = x.Die1.T0_Result;
        //            sheet.Cells["BO" + rowId.ToString()].Value = x.Die1.T0_Solve_Method;
        //            sheet.Cells["BP" + rowId.ToString()].Value = x.Die1.T0_Solve_Result;
        //            sheet.Cells["BQ" + rowId.ToString()].Value = x.Die1.Texture_Meeting_Date;
        //            sheet.Cells["BR" + rowId.ToString()].Value = x.Die1.Texture_Go_Date;
        //            sheet.Cells["BS" + rowId.ToString()].Value = x.Die1.S0_Plan.HasValue ? x.Die1.S0_Plan.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["BT" + rowId.ToString()].Value = x.Die1.S0_Result;
        //            sheet.Cells["BU" + rowId.ToString()].Value = x.Die1.S0_Solve_Method;
        //            sheet.Cells["BV" + rowId.ToString()].Value = x.Die1.S0_solve_Result;
        //            sheet.Cells["BW" + rowId.ToString()].Value = x.Die1.Texture_App_Date;
        //            sheet.Cells["BX" + rowId.ToString()].Value = x.Die1.Texture_Internal_App_Result;
        //            sheet.Cells["BY" + rowId.ToString()].Value = x.Die1.Texture_JP_HP_App_Result;
        //            sheet.Cells["BZ" + rowId.ToString()].Value = x.Die1.Texture_Note;
        //            sheet.Cells["CA" + rowId.ToString()].Value = x.Die1.PreKK_Plan;
        //            sheet.Cells["CB" + rowId.ToString()].Value = x.Die1.PreKK_Actual;
        //            sheet.Cells["CC" + rowId.ToString()].Value = x.Die1.PreKK_Result;
        //            sheet.Cells["CD" + rowId.ToString()].Value = x.Die1.FA_Sub_Time;
        //            sheet.Cells["CE" + rowId.ToString()].Value = x.Die1.FA_Plan.HasValue ? x.Die1.FA_Plan.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["CF" + rowId.ToString()].Value = x.Die1.FA_Result_Date.HasValue ? x.Die1.FA_Result_Date.Value.ToString("MM/dd/yyyy") : "";
        //            sheet.Cells["CG" + rowId.ToString()].Value = x.Die1.FA_Result;
        //            sheet.Cells["CH" + rowId.ToString()].Value = x.Die1.FA_Problem;
        //            sheet.Cells["CI" + rowId.ToString()].Value = x.Die1.FA_Action_Improve;
        //            sheet.Cells["CJ" + rowId.ToString()].Value = x.Die1.MT1_Date;
        //            sheet.Cells["CK" + rowId.ToString()].Value = x.Die1.MT1_Gather_Date;
        //            sheet.Cells["CL" + rowId.ToString()].Value = x.Die1.MT1_Problem;
        //            sheet.Cells["CM" + rowId.ToString()].Value = x.Die1.MT1_Remark;
        //            sheet.Cells["CN" + rowId.ToString()].Value = x.Die1.MTF_Date;
        //            sheet.Cells["CO" + rowId.ToString()].Value = x.Die1.MTF_Gather_Date;
        //            sheet.Cells["CP" + rowId.ToString()].Value = x.Die1.MTF_Problem;
        //            sheet.Cells["CQ" + rowId.ToString()].Value = x.Die1.MTF_Remark;
        //            sheet.Cells["CR" + rowId.ToString()].Value = x.Die1.First_Lot_Date;
        //            sheet.Cells["CS" + rowId.ToString()].Value = x.Die1.DieCost_USD;
        //            sheet.Cells["CT" + rowId.ToString()].Value = x.Die1.DieCost_JPY;
        //            sheet.Cells["CU" + rowId.ToString()].Value = x.Die1.DieCost_VND;
        //            sheet.Cells["CV" + rowId.ToString()].Value = x.Die1.DieWarranty_Short;
        //            sheet.Cells["CW" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.RecodeDate;
        //            sheet.Cells["CX" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.ActualShort;
        //            sheet.Cells["CY" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.UsingStatus;
        //            sheet.Cells["CZ" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.DetailUsingStatus;
        //            sheet.Cells["DA" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.DieOperationStatus;
        //            sheet.Cells["DB" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.StopDate;
        //            sheet.Cells["DC" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.SupplierProposal;
        //            sheet.Cells["DD" + rowId.ToString()].Value = x.Die1.DieStatusUpdateRegulars.LastOrDefault()?.ReasonProposal;
        //            sheet.Cells["DE" + rowId.ToString()].Value = db.DieLendingRequests.Where(trs => trs.DieNo == x.DieNo && trs.Active != false && trs.LendingStatusID == 10).OrderByDescending(trs => trs.LendingRequestID).FirstOrDefault()?.CurrentLocation;
        //            sheet.Cells["DF" + rowId.ToString()].Value = db.DieLendingRequests.Where(trs => trs.DieNo == x.DieNo && trs.Active != false && trs.LendingStatusID == 10).OrderByDescending(trs => trs.LendingRequestID).FirstOrDefault()?.NewLocation;
        //            sheet.Cells["DG" + rowId.ToString()].Value = x.Die1.EditBy;
        //            sheet.Cells["DH" + rowId.ToString()].Value = x.Die1.EditDate;
        //            sheet.Cells["DI" + rowId.ToString()].Value = x.Die1.His_Update;
        //            sheet.Cells["DJ" + rowId.ToString()].Value = x.Die1.IssueDate;
        //            sheet.Cells["DK" + rowId.ToString()].Value = x.Die1.isOfficial;
        //            sheet.Cells["DL" + rowId.ToString()].Value = x.Die1.isCancel;
        //            sheet.Cells["DM" + rowId.ToString()].Value = x.Die1.FixedAssetNo;
        //            sheet.Cells["DN" + rowId.ToString()].Value = x.Die1.Belong;

        //            i++;
        //            rowId++;
        //        }
        //        package.SaveAs(output);
        //        output.Position = 0;
        //        TempData[handle] = output.ToArray();
        //    }

        //    var data = new { FileGuid = handle, FileName = DateTime.Now.ToString("yyyyMMdd-HHmmss") + "_DieControlList.xlsx" };
        //    return data;
        //}

        public Object exportToControlList(List<CommonDie1> records)
        {
            int userID = int.Parse(Session["UserID"].ToString());
            string handle = Guid.NewGuid().ToString();
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/UpdateDieInfo/Format_Die_Master_List02.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();
                int rowId = 9;
                int i = 1;
                sheet.Cells["A3"].Value = "Date: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                foreach (var x in records)
                {
                    var pend = commonFunction.getPendingAndDeptResponse(x.Die1);
                    sheet.Cells["A" + rowId.ToString()].Value = i;
                    sheet.Cells["B" + rowId.ToString()].Value = x.Die1.FixedAssetNo;
                    sheet.Cells["C" + rowId.ToString()].Value = x.Die1.DieClassify;
                    sheet.Cells["D" + rowId.ToString()].Value = x.Die1.Step;
                    sheet.Cells["E" + rowId.ToString()].Value = x.Die1.Rank;
                    sheet.Cells["F" + rowId.ToString()].Value = x.PartNo;
                    sheet.Cells["G" + rowId.ToString()].Value = x.Parts1.PartName;
                    sheet.Cells["H" + rowId.ToString()].Value = x.Die1.ProcessCodeCalogory.Type;
                    sheet.Cells["I" + rowId.ToString()].Value = x.Die1.Die_Code;
                    sheet.Cells["J" + rowId.ToString()].Value = x.Die1.DieNo;
                    sheet.Cells["K" + rowId.ToString()].Value = x.Die1.ModelID > 0 ? x.Die1.ModelList.ModelName : "";
                    sheet.Cells["L" + rowId.ToString()].Value = x.Die1.SupplierID >= 0 ? x.Die1.Supplier.SupplierCode : "";
                    sheet.Cells["M" + rowId.ToString()].Value = x.Die1.SupplierID >= 0 ? x.Die1.Supplier.SupplierName : "";
                    sheet.Cells["N" + rowId.ToString()].Value = x.Die1.DieMaker;
                    sheet.Cells["O" + rowId.ToString()].Value = x.Die1.DieMakeLocation;
                    sheet.Cells["P" + rowId.ToString()].Value = x.Die1.MCsize;
                    sheet.Cells["Q" + rowId.ToString()].Value = x.Die1.CavQuantity;
                    sheet.Cells["R" + rowId.ToString()].Value = x.Die1.CycleTime_Sec;
                    sheet.Cells["S" + rowId.ToString()].Value = x.Die1.CycleTime_TargetAsDSUM;
                    sheet.Cells["T" + rowId.ToString()].Value = x.Die1.SPM_Progressive;
                    sheet.Cells["U" + rowId.ToString()].Value = x.Die1.SPM_Single;
                    sheet.Cells["V" + rowId.ToString()].Value = x.Die1.Common_Part_With;
                    sheet.Cells["W" + rowId.ToString()].Value = x.Die1.Family_Die_With;
                    sheet.Cells["X" + rowId.ToString()].Value = pend.Progress;
                    sheet.Cells["Y" + rowId.ToString()].Value = commonFunction.getStatus(x.Die1);
                    sheet.Cells["X" + rowId.ToString()].Value = pend.Pending_Status;
                    sheet.Cells["AA" + rowId.ToString()].Value = pend.Dept_Respone;
                    sheet.Cells["AB" + rowId.ToString()].Value = commonFunction.getWarning(x.Die1);
                    sheet.Cells["AC" + rowId.ToString()].Value = x.Die1.Genaral_Information;
                    sheet.Cells["AD" + rowId.ToString()].Value = x.Die1.Decision_Date;
                    sheet.Cells["AE" + rowId.ToString()].Value = x.Die1.Select_Supplier_Date;
                    sheet.Cells["AF" + rowId.ToString()].Value = x.Die1.QTN_Sub_Date;
                    sheet.Cells["AG" + rowId.ToString()].Value = x.Die1.QTN_App_Date;
                    sheet.Cells["AH" + rowId.ToString()].Value = x.Die1.Need_Use_Date;
                    sheet.Cells["AI" + rowId.ToString()].Value = x.Die1.Target_OK_Date;
                    sheet.Cells["AJ" + rowId.ToString()].Value = x.Die1.Inv_Idea;
                    sheet.Cells["AK" + rowId.ToString()].Value = x.Die1.Inv_FB_To;
                    sheet.Cells["AL" + rowId.ToString()].Value = x.Die1.Inv_Result;
                    sheet.Cells["AM" + rowId.ToString()].Value = x.Die1.Inv_Cost_Down;
                    sheet.Cells["AN" + rowId.ToString()].Value = x.Die1.DFM_Sub_Date;
                    sheet.Cells["AO" + rowId.ToString()].Value = x.Die1.DFM_PAE_Check_Date;
                    sheet.Cells["AP" + rowId.ToString()].Value = x.Die1.DFMPE1Checked;
                    sheet.Cells["AQ" + rowId.ToString()].Value = x.Die1.DFM_PE_App_Date;
                    sheet.Cells["AR" + rowId.ToString()].Value = x.Die1.DFM_PAE_App_Date;
                    sheet.Cells["AS" + rowId.ToString()].Value = x.Die1.CoreCavMaterial;
                    sheet.Cells["AT" + rowId.ToString()].Value = x.Die1.SliderMaterial;
                    sheet.Cells["AU" + rowId.ToString()].Value = x.Die1.LifterMaterial;
                    sheet.Cells["AV" + rowId.ToString()].Value = x.Die1.PunchBackingPlate;
                    sheet.Cells["AW" + rowId.ToString()].Value = x.Die1.PunchPlate;
                    sheet.Cells["AX" + rowId.ToString()].Value = x.Die1.StripperBackingPlate;
                    sheet.Cells["AY" + rowId.ToString()].Value = x.Die1.StripperPlate;
                    sheet.Cells["AZ" + rowId.ToString()].Value = x.Die1.DiePlate;
                    sheet.Cells["BA" + rowId.ToString()].Value = x.Die1.DieBackingPlate;
                    sheet.Cells["BB" + rowId.ToString()].Value = x.Die1.Punch;
                    sheet.Cells["BC" + rowId.ToString()].Value = x.Die1.InsertBlock;
                    sheet.Cells["BD" + rowId.ToString()].Value = x.Die1.isStacking == true ? "Yes" : x.Die1.isStacking == false ? "No" : "";
                    sheet.Cells["BE" + rowId.ToString()].Value = x.Die1.isShaftKashime == true ? "Yes" : x.Die1.isShaftKashime == false ? "No" : "";
                    sheet.Cells["BF" + rowId.ToString()].Value = x.Die1.isBurringKashime == true ? "Yes" : x.Die1.isBurringKashime == false ? "No" : "";
                    sheet.Cells["BG" + rowId.ToString()].Value = x.Die1.PXNoOfComponent;
                    sheet.Cells["BH" + rowId.ToString()].Value = x.Die1.HotRunner;
                    sheet.Cells["BI" + rowId.ToString()].Value = x.Die1.GateType;
                    sheet.Cells["BJ" + rowId.ToString()].Value = x.Die1.SpecialSpec;
                    sheet.Cells["BK" + rowId.ToString()].Value = x.Die1.DSUM_Idea;
                    sheet.Cells["BL" + rowId.ToString()].Value = x.Die1.NOofForQ;
                    sheet.Cells["BM" + rowId.ToString()].Value = x.Die1.NOofForC;
                    sheet.Cells["BN" + rowId.ToString()].Value = x.Die1.NOofForC;

                    sheet.Cells["BO" + rowId.ToString()].Value = x.Die1.MR_Request_Date.HasValue ? x.Die1.MR_Request_Date.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["BP" + rowId.ToString()].Value = x.Die1.MR_App_Date.HasValue ? x.Die1.MR_App_Date.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["BQ" + rowId.ToString()].Value = x.Die1.PO_Issue_Date.HasValue ? x.Die1.PO_Issue_Date.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["BR" + rowId.ToString()].Value = x.Die1.PODate.HasValue ? x.Die1.PODate.Value.ToString("MM/dd/yyyy") : "";

                    sheet.Cells["BS" + rowId.ToString()].Value = x.Die1.Design_Check_Plan;
                    sheet.Cells["BT" + rowId.ToString()].Value = x.Die1.Design_Check_Actual;
                    sheet.Cells["BU" + rowId.ToString()].Value = x.Die1.Design_Check_Result;
                    sheet.Cells["BV" + rowId.ToString()].Value = x.Die1.NoOfPoit_Not_FL_DMF;
                    sheet.Cells["BW" + rowId.ToString()].Value = x.Die1.NoOfPoint_Not_FL_Spec;
                    sheet.Cells["BX" + rowId.ToString()].Value = x.Die1.NoOfPoint_Kaizen;

                    sheet.Cells["BY" + rowId.ToString()].Value = x.Die1.JIG_Using == true ? "Yes" : x.Die1.JIG_Using == false ? "No" : ""; ;
                    sheet.Cells["BZ" + rowId.ToString()].Value = x.Die1.JIG_No;
                    sheet.Cells["CA" + rowId.ToString()].Value = x.Die1.JIG_Check_Plan;
                    sheet.Cells["CB" + rowId.ToString()].Value = x.Die1.JIG_Check_Result;
                    sheet.Cells["CC" + rowId.ToString()].Value = x.Die1.JIG_ETA_Supplier;
                    sheet.Cells["CD" + rowId.ToString()].Value = x.Die1.JIG_Status;

                    sheet.Cells["CE" + rowId.ToString()].Value = x.Die1.T0_Plan.HasValue ? x.Die1.T0_Plan.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["CF" + rowId.ToString()].Value = x.Die1.T0_Actual.HasValue ? x.Die1.T0_Actual.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["CG" + rowId.ToString()].Value = x.Die1.T0_Result;
                    sheet.Cells["CH" + rowId.ToString()].Value = x.Die1.T0_Solve_Method;
                    sheet.Cells["CI" + rowId.ToString()].Value = x.Die1.T0_Solve_Result;

                    sheet.Cells["CJ" + rowId.ToString()].Value = x.Die1.Texture == true ? "Yes" : x.Die1.Texture == false ? "No" : ""; ;
                    sheet.Cells["CK" + rowId.ToString()].Value = x.Die1.Texture_Meeting_Date;
                    sheet.Cells["CL" + rowId.ToString()].Value = x.Die1.Texture_Go_Date;
                    sheet.Cells["CM" + rowId.ToString()].Value = x.Die1.S0_Plan.HasValue ? x.Die1.S0_Plan.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["CN" + rowId.ToString()].Value = x.Die1.S0_Result;
                    sheet.Cells["CO" + rowId.ToString()].Value = x.Die1.S0_Solve_Method;
                    sheet.Cells["CP" + rowId.ToString()].Value = x.Die1.S0_solve_Result;
                    sheet.Cells["CQ" + rowId.ToString()].Value = x.Die1.Texture_App_Date;
                    sheet.Cells["CR" + rowId.ToString()].Value = x.Die1.Texture_Internal_App_Result;
                    sheet.Cells["CS" + rowId.ToString()].Value = x.Die1.Texture_JP_HP_App_Result;
                    sheet.Cells["CT" + rowId.ToString()].Value = x.Die1.Texture_Note;

                    sheet.Cells["CU" + rowId.ToString()].Value = x.Die1.PreKK_Plan;
                    sheet.Cells["CV" + rowId.ToString()].Value = x.Die1.PreKK_Actual;
                    sheet.Cells["CW" + rowId.ToString()].Value = x.Die1.PreKK_Result;

                    sheet.Cells["CX" + rowId.ToString()].Value = x.Die1.FA_Sub_Time;
                    sheet.Cells["CY" + rowId.ToString()].Value = x.Die1.FA_Plan.HasValue ? x.Die1.FA_Plan.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["CZ" + rowId.ToString()].Value = x.Die1.FA_Result_Date.HasValue ? x.Die1.FA_Result_Date.Value.ToString("MM/dd/yyyy") : "";
                    sheet.Cells["DA" + rowId.ToString()].Value = x.Die1.FA_Result;
                    sheet.Cells["DB" + rowId.ToString()].Value = x.Die1.FA_Problem;
                    sheet.Cells["DC" + rowId.ToString()].Value = x.Die1.FA_Action_Improve;

                    sheet.Cells["DD" + rowId.ToString()].Value = x.Die1.MT1_Date;
                    sheet.Cells["DE" + rowId.ToString()].Value = x.Die1.MT1_Gather_Date;
                    sheet.Cells["DF" + rowId.ToString()].Value = x.Die1.MT1_Problem;
                    sheet.Cells["DG" + rowId.ToString()].Value = x.Die1.MT1_Remark;
                    sheet.Cells["DH" + rowId.ToString()].Value = x.Die1.MTF_Date;
                    sheet.Cells["DI" + rowId.ToString()].Value = x.Die1.MTF_Gather_Date;
                    sheet.Cells["DJ" + rowId.ToString()].Value = x.Die1.MTF_Problem;
                    sheet.Cells["DK" + rowId.ToString()].Value = x.Die1.MTF_Remark;

                    sheet.Cells["DL" + rowId.ToString()].Value = x.Die1.First_Lot_Date;
                    sheet.Cells["DM" + rowId.ToString()].Value = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_USD?.ToString() : "***";
                    sheet.Cells["DN" + rowId.ToString()].Value = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_JPY?.ToString() : "***";
                    sheet.Cells["DO" + rowId.ToString()].Value = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_JPY?.ToString() : "***";
                    sheet.Cells["DP" + rowId.ToString()].Value = x.Die1.DieWarranty_Short;
                    sheet.Cells["DQ" + rowId.ToString()].Value = x.Die1.WarrantyShotAsDSUM;
                    sheet.Cells["DR" + rowId.ToString()].Value = x.Die1.RecordDate;
                    sheet.Cells["DS" + rowId.ToString()].Value = x.Die1.Short;
                    sheet.Cells["DT" + rowId.ToString()].Value = x.Die1.DieStatusCategory?.Type;
                    sheet.Cells["DU" + rowId.ToString()].Value = x.Die1.InventoryStatus;
                    sheet.Cells["DV" + rowId.ToString()].Value = x.Die1.RemarkDieStatusUsing;
                    sheet.Cells["DW" + rowId.ToString()].Value = x.Die1.StopDate;
                    sheet.Cells["DX" + rowId.ToString()].Value = x.Die1.SupplierProposal;
                    sheet.Cells["DY" + rowId.ToString()].Value = x.Die1.ReasonProposal;
                    sheet.Cells["DZ" + rowId.ToString()].Value = db.DieLendingRequests.Where(trs => trs.DieNo == x.DieNo && trs.Active != false && trs.LendingStatusID == 10).OrderByDescending(trs => trs.LendingRequestID).FirstOrDefault()?.CurrentLocation;
                    sheet.Cells["EA" + rowId.ToString()].Value = db.DieLendingRequests.Where(trs => trs.DieNo == x.DieNo && trs.Active != false && trs.LendingStatusID == 10).OrderByDescending(trs => trs.LendingRequestID).FirstOrDefault()?.NewLocation;

                    sheet.Cells["EB" + rowId.ToString()].Value = x.Die1.EditBy;
                    sheet.Cells["EC" + rowId.ToString()].Value = x.Die1.EditDate;
                    sheet.Cells["ED" + rowId.ToString()].Value = x.Die1.His_Update;
                    sheet.Cells["EE" + rowId.ToString()].Value = x.Die1.IssueDate;
                    sheet.Cells["EF" + rowId.ToString()].Value = x.Die1.isOfficial;
                    sheet.Cells["EG" + rowId.ToString()].Value = x.Die1.isCancel;
                    sheet.Cells["EH" + rowId.ToString()].Value = x.Die1.Belong;

                    i++;
                    rowId++;
                }
                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }

            var data = new { FileGuid = handle, FileName = DateTime.Now.ToString("yyyyMMdd-HHmmss") + "_DieControlList.xlsx" };
            return data;
        }

        public JsonResult importControlList(HttpPostedFileBase file)
        {
            // Form dùng chung
            // Phòng nào thì chỉ cho import các trường của phòng đó.

            var today = DateTime.Now;
            var dept = Session["Dept"].ToString().Trim();
            var role = Session["Die_Lauch_Role"].ToString().Trim();
            bool admin = Session["Admin"].ToString() == "Admin";
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            if (role == "Edit" || admin)
            {
                // Doc file
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    // Check form
                    var A8 = worksheet.Cells["A8"].Text; // No
                    var F8 = worksheet.Cells["F8"].Text; // Part_No
                    var I8 = worksheet.Cells["I8"].Text; // Die_No
                    var EH8 = worksheet.Cells["EH8"].Text; //Refer

                    if (A8 != "No" || F8 != "Part_No" || I8 != "Die_No" || EH8 != "Belong")
                    {
                        fail.Add("Sai Format, Format đã bị thêm cột");
                        var data1 = new
                        {
                            success = success.Count(),
                            fail = fail
                        };
                        return Json(data1, JsonRequestBehavior.AllowGet);
                    }
                    // Kết thúc check form

                    for (int row = start.Row + 8; row <= end.Row; row++)
                    { // Row by row...
                        var partNo = worksheet.Cells[row, 6].Text;
                        if (String.IsNullOrEmpty(partNo)) break;
                        var dieCode = worksheet.Cells[row, 9].Text;
                        var findRecord = db.Die1.Where(x => x.PartNoOriginal == partNo && x.Die_Code == dieCode && x.Active != false).FirstOrDefault();

                        if (findRecord == null)
                        {
                            goto exitLoop;
                        }

                        // Admin duoc sua cot nao?
                        if (admin == true)
                        {
                            // Verify modelName
                            var modelName = worksheet.Cells[row, 11].Text.Trim().ToUpper();
                            var model = db.ModelLists.Where(x => x.ModelName.ToUpper().Contains(modelName)).FirstOrDefault();
                            if (model == null)
                            {
                                fail.Add(partNo + "-" + dieCode + ": Lỗi ko tồn tại Model " + modelName);
                            }
                            else
                            {
                                findRecord.ModelID = model.ModelID;
                            }


                        }



                        //3. MQA
                        if (dept.Contains("MQA") || admin)
                        {
                            var isJIG = worksheet.Cells[row, 77].Text.Trim().ToUpper();
                            if (!String.IsNullOrEmpty(isJIG))
                            {
                                findRecord.JIG_Using = (isJIG == "Yes" || isJIG == "Y" || isJIG == "y" || isJIG == "have" || isJIG == "Have") ? true : false;
                            }

                            findRecord.JIG_No = worksheet.Cells[row, 78].Text.Trim().ToUpper();
                            findRecord.JIG_Check_Plan = worksheet.Cells[row, 79].Text.Trim().ToUpper();
                            findRecord.JIG_Check_Result = worksheet.Cells[row, 80].Text.Trim().ToUpper();
                            findRecord.JIG_ETA_Supplier = worksheet.Cells[row, 81].Text.Trim().ToUpper();
                            findRecord.JIG_Status = worksheet.Cells[row, 82].Text.Trim().ToUpper();
                        }


                        // 4. PAE
                        if (dept.Contains("PAE") || admin)
                        {
                            findRecord.Step = worksheet.Cells[row, 4].Text.Trim().ToUpper();
                            findRecord.Rank = worksheet.Cells[row, 5].Text.Trim().ToUpper();

                            try
                            {
                                findRecord.Target_OK_Date = DateTime.Parse(worksheet.Cells[row, 35].Text.Trim().ToUpper());
                            }
                            catch { }
                            //findRecord.Inv_Idea = worksheet.Cells[row, 25].Text.Trim().ToUpper();
                            //findRecord.Inv_FB_To = worksheet.Cells[row, 26].Text.Trim().ToUpper();
                            //findRecord.Inv_Result = worksheet.Cells[row, 27].Text.Trim().ToUpper();
                            //findRecord.Inv_Cost_Down = worksheet.Cells[row, 28].Text.Trim().ToUpper();
                            //findRecord.CoreCavMaterial = worksheet.Cells[row, 34].Text.Trim().ToUpper();
                            //findRecord.SliderMaterial = worksheet.Cells[row, 35].Text.Trim().ToUpper();
                            //findRecord.LifterMaterial = worksheet.Cells[row, 36].Text.Trim().ToUpper();
                            //var isTexture = worksheet.Cells[row, 37].Text.Trim().ToUpper();
                            //if (!String.IsNullOrEmpty(isTexture))
                            //{
                            //    findRecord.Texture = (isTexture == "YES" || isTexture == "Y" || isTexture == "1" || isTexture == "HAVE" || isTexture == "TRUE") ? true : false;
                            //}
                            //findRecord.HotRunner = worksheet.Cells[row, 38].Text.Trim().ToUpper();
                            //findRecord.GateType = worksheet.Cells[row, 39].Text.Trim().ToUpper();
                            //try
                            //{
                            //    findRecord.MCsize = int.Parse(worksheet.Cells[row, 40].Text.Trim().ToUpper());
                            //    findRecord.CavQuantity = int.Parse(worksheet.Cells[row, 41].Text.Trim().ToUpper());

                            //}
                            //catch
                            //{
                            //    fail.Add(partNo + "-" + dieCode + ": MC size va Cav_Qty Phải là số ");
                            //}

                            //findRecord.DieMakeLocation = worksheet.Cells[row, 42].Text.Trim().ToUpper();
                            //findRecord.DieMaker = worksheet.Cells[row, 43].Text.Trim().ToUpper();
                            //findRecord.Family_Die_With = worksheet.Cells[row, 44].Text.Trim().ToUpper();
                            //findRecord.Common_Part_With = worksheet.Cells[row, 45].Text.Trim().ToUpper();
                            //findRecord.SpecialSpec = worksheet.Cells[row, 46].Text.Trim().ToUpper();
                            //findRecord.DSUM_Idea = worksheet.Cells[row, 47].Text.Trim().ToUpper();


                            //findRecord.Design_Check_Plan = worksheet.Cells[row, 71].Text.Trim().ToUpper();
                            //findRecord.Design_Check_Actual = worksheet.Cells[row, 72].Text.Trim().ToUpper();
                            //findRecord.Design_Check_Result = worksheet.Cells[row, 73].Text.Trim().ToUpper();
                            //findRecord.NoOfPoit_Not_FL_DMF = worksheet.Cells[row, 74].Text.Trim().ToUpper();
                            //findRecord.NoOfPoint_Not_FL_Spec = worksheet.Cells[row, 75].Text.Trim().ToUpper();
                            //findRecord.NoOfPoint_Kaizen = worksheet.Cells[row, 76].Text.Trim().ToUpper();


                            //findRecord.T0_Result = worksheet.Cells[row, 85].Text.Trim().ToUpper();
                            //findRecord.T0_Solve_Method = worksheet.Cells[row, 86].Text.Trim().ToUpper();
                            //findRecord.T0_Solve_Result = worksheet.Cells[row, 87].Text.Trim().ToUpper();
                            //findRecord.Texture_Meeting_Date = worksheet.Cells[row, 69].Text.Trim().ToUpper();
                            //findRecord.Texture_Go_Date = worksheet.Cells[row, 70].Text.Trim().ToUpper();
                            //try
                            //{
                            //    var soPlan = worksheet.Cells[row, 71].Text.Trim().ToUpper();
                            //    if (!String.IsNullOrWhiteSpace(soPlan))
                            //    {
                            //        findRecord.S0_Plan = DateTime.Parse(worksheet.Cells[row, 71].Text.Trim().ToUpper());
                            //    }
                            //}
                            //catch
                            //{
                            //    fail.Add(partNo + "-" + dieCode + ": Lỗi Sai S0_Plan format Date MM/DD/YYYY ");
                            //}

                            //findRecord.S0_Result = worksheet.Cells[row, 72].Text.Trim().ToUpper();
                            //findRecord.S0_Solve_Method = worksheet.Cells[row, 73].Text.Trim().ToUpper();
                            //findRecord.S0_solve_Result = worksheet.Cells[row, 74].Text.Trim().ToUpper();
                            //findRecord.Texture_App_Date = worksheet.Cells[row, 75].Text.Trim().ToUpper();
                            //findRecord.Texture_Internal_App_Result = worksheet.Cells[row, 76].Text.Trim().ToUpper();
                            //findRecord.Texture_JP_HP_App_Result = worksheet.Cells[row, 77].Text.Trim().ToUpper();
                            //findRecord.Texture_Note = worksheet.Cells[row, 78].Text.Trim().ToUpper();


                        }



                        //// 5. PAE & PUR
                        //if (dept.Contains("PAE") || dept.Contains("PUR") || admin)
                        //{

                        //    findRecord.Decision_Date = worksheet.Cells[row, 19].Text.Trim().ToUpper();
                        //    findRecord.MT1_Date = worksheet.Cells[row, 88].Text.Trim().ToUpper();
                        //    findRecord.MT1_Gather_Date = worksheet.Cells[row, 89].Text.Trim().ToUpper();
                        //    findRecord.MTF_Date = worksheet.Cells[row, 92].Text.Trim().ToUpper();
                        //    findRecord.MTF_Gather_Date = worksheet.Cells[row, 93].Text.Trim().ToUpper();
                        //}


                        // 6. 
                        if (dept.Contains("PE1") || admin)
                        {
                            try
                            {
                                var faRsDate = worksheet.Cells[row, 104].Text.Trim().ToUpper();
                                if (!String.IsNullOrWhiteSpace(faRsDate))
                                {
                                    findRecord.FA_Result_Date = DateTime.Parse(faRsDate);
                                }
                            }
                            catch
                            {
                                //fail.Add(partNo + "-" + dieCode + ": Lỗi FA_Result_Date format Date MM/DD/YYYY ");
                            }
                            findRecord.FA_Result = worksheet.Cells[row, 105].Text.Trim().ToUpper();

                        }

                        if (dept == "PUR" || admin)
                        {
                           // findRecord.QTN_Sub_Date = worksheet.Cells[row, 21].Text.Trim().ToUpper();
                            findRecord.Need_Use_Date = worksheet.Cells[row, 34].Text.Trim().ToUpper();

                            try
                            {
                                var toPlan = worksheet.Cells[row, 83].Text.Trim().ToUpper();
                                if (!String.IsNullOrWhiteSpace(toPlan))
                                {
                                    findRecord.T0_Plan = DateTime.Parse(toPlan);
                                }
                            }
                            catch
                            {
                               // fail.Add(partNo + "-" + dieCode + ": Lỗi T0_Plan format Date MM/DD/YYYY ");
                            }
                            try
                            {
                                var toactual = worksheet.Cells[row, 84].Text.Trim().ToUpper();
                                if (!String.IsNullOrWhiteSpace(toactual))
                                {
                                    findRecord.T0_Actual = DateTime.Parse(toactual);
                                }
                            }
                            catch
                            {
                               // fail.Add(partNo + "-" + dieCode + ": Lỗi T0_Actual format Date MM/DD/YYYY ");
                            }


                            findRecord.PreKK_Plan = worksheet.Cells[row, 99].Text.Trim().ToUpper();
                            findRecord.PreKK_Actual = worksheet.Cells[row, 100].Text.Trim().ToUpper();
                            findRecord.PreKK_Result = worksheet.Cells[row, 101].Text.Trim().ToUpper();
                            try
                            {
                                var faPlan = worksheet.Cells[row, 103].Text.Trim().ToUpper();
                                if (!String.IsNullOrWhiteSpace(faPlan))
                                {
                                    findRecord.FA_Plan = DateTime.Parse(faPlan);
                                }
                            }
                            catch
                            {
                                fail.Add(partNo + "-" + dieCode + ": Lỗi FA_Plan format Date MM/DD/YYYY ");
                            }
                            findRecord.First_Lot_Date = worksheet.Cells[row, 116].Text.Trim().ToUpper();

                        }


                        //if (dept == "PUS" || admin)
                        //{
                        //    findRecord.Select_Supplier_Date = worksheet.Cells[row, 20].Text.Trim().ToUpper();
                        //    findRecord.QTN_App_Date = worksheet.Cells[row, 22].Text.Trim().ToUpper();
                        //    // Verify SupplierCode.
                        //    var supplierCode = worksheet.Cells[row, 11].Text.Trim().ToUpper();
                        //    var supplier = db.Suppliers.Where(x => x.SupplierCode == supplierCode).FirstOrDefault();
                        //    if (supplier == null)
                        //    {
                        //        fail.Add(partNo + "-" + dieCode + ": Lỗi ko tồn tại Supplier Code " + supplierCode);
                        //    }
                        //    else
                        //    {
                        //        findRecord.SupplierID = supplier.SupplierID;
                        //    }
                        //}

                        //auto 
                        findRecord.EditBy = Session["Name"].ToString();
                        findRecord.EditDate = today;
                        findRecord.His_Update = today.ToString("yy-MM-dd hh-mm") + Session["Name"].ToString() + " updated by excel file " + System.Environment.NewLine + findRecord.His_Update;
                        // Luu 
                        // findRecord = commonFunction.updateDieStatus(findRecord);
                        db.Entry(findRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success.Add(partNo + "-" + dieCode);

                    exitLoop:
                        ViewBag.msg = "just for exit loop";
                    }
                }
            }
            var data = new
            {
                success = success.Count(),
                fail = fail
            };
            return Json(data, JsonRequestBehavior.AllowGet);

        }
        public Object dataReturnView(List<CommonDie1> records)
        {
            int userID = int.Parse(Session["UserID"].ToString());
            List<object> outPut = new List<object>();
            foreach (CommonDie1 x in records)
            {
                var pend = commonFunction.getPendingAndDeptResponse(x.Die1);
                outPut.Add(new
                {
                    DieID = x.Die1.DieID,
                    Step = x.Die1.Step,
                    Rank = x.Die1.Rank,
                    Category = x.Die1.DieClassify,
                    PartNo = x.PartNo,
                    PartName = x.Parts1.PartName,
                    Process_Code = x.Die1.ProcessCodeCalogory.Type,
                    Die_Code = x.Die1.Die_Code,
                    FixedAssetNo = x.Die1.FixedAssetNo,
                    DieNo = x.Die1.DieNo,
                    ModelID = x.Die1.ModelID,
                    Model_Name = x.Die1.ModelID > 0 ? db.ModelLists.Find(x.Die1.ModelID).ModelName : "Invalid",
                    SupplierID = x.Die1.SupplierID,
                    Supplier_Name = x.Die1.SupplierID >= 0 ? db.Suppliers.Find(x.Die1.SupplierID).SupplierName : "Invalid",
                    Supplier_Code = x.Die1.SupplierID >= 0 ? db.Suppliers.Find(x.Die1.SupplierID).SupplierCode : "Invalid",
                    CycleTime_Sec = x.Die1.CycleTime_Sec,
                    CycleTime_TargetAsDSUM = x.Die1.CycleTime_TargetAsDSUM,
                    SPM_Progressive = x.Die1.SPM_Progressive,
                    SPM_Single = x.Die1.SPM_Single,
                    Texture = x.Die1.Texture == null ? null : x.Die1.Texture == true ? "Yes" : "No",
                    Status = commonFunction.getStatus(x.Die1),
                    Pending_Status = pend.Pending_Status,
                    Dept_Responsibility = pend.Dept_Respone,
                    Progress = pend.Progress,
                    Warning = commonFunction.getWarning(x.Die1),
                    Genaral_Information = x.Die1.Genaral_Information,
                    Decision_Date = x.Die1.Decision_Date,
                    Select_Supplier_Date = x.Die1.Select_Supplier_Date,
                    QTN_Sub_Date = x.Die1.QTN_Sub_Date,
                    QTN_App_Date = x.Die1.QTN_App_Date,
                    Need_Use_Date = x.Die1.Need_Use_Date,
                    Target_OK_Date = x.Die1.Target_OK_Date.HasValue ? x.Die1.Target_OK_Date.Value.ToString("MM/dd/yyyy") : null,
                    Inv_Idea = x.Die1.Inv_Idea,
                    Inv_FB_To = x.Die1.Inv_FB_To,
                    Inv_Result = x.Die1.Inv_Result,
                    Inv_Cost_Down = x.Die1.Inv_Cost_Down,
                    DFM_Sub_Date = x.Die1.DFM_Sub_Date,
                    DFM_PAE_Check_Date = x.Die1.DFM_PAE_Check_Date,
                    DFM_PE_Check_Date = x.Die1.DFM_PE_Check_Date,
                    DFM_PE_App_Date = x.Die1.DFM_PE_App_Date,
                    DFM_PAE_App_Date = x.Die1.DFM_PAE_App_Date,
                    DFMMaker = x.Die1.AttDFMMaker,
                    DFMPAEChecked = x.Die1.AttDFMPAEChecked,
                    DFMPE1Checked = x.Die1.AttDFMPE1Checked,
                    DFMPE1App = x.Die1.AttDFMPE1App,
                    DFMPAEApp = x.Die1.AttDFMPAEApp,
                    CoreCavMaterial = x.Die1.CoreCavMaterial,
                    SliderMaterial = x.Die1.SliderMaterial,
                    LifterMaterial = x.Die1.LifterMaterial,

                    PunchBackingPlate = x.Die1.PunchBackingPlate,
                    PunchPlate = x.Die1.PunchPlate,
                    StripperBackingPlate = x.Die1.StripperBackingPlate,
                    StripperPlate = x.Die1.StripperPlate,
                    DiePlate = x.Die1.DiePlate,
                    DieBackingPlate = x.Die1.DieBackingPlate,
                    InsertBlock = x.Die1.InsertBlock,
                    Punch = x.Die1.Punch,
                    isStacking = x.Die1.isStacking == null ? null : x.Die1.isStacking == true ? "Yes" : "No",
                    isShaftKashime = x.Die1.isShaftKashime == null ? null : x.Die1.isShaftKashime == true ? "Yes" : "No",
                    isBurringKashime = x.Die1.isBurringKashime == null ? null : x.Die1.isBurringKashime == true ? "Yes" : "No",
                    PXNoOfComponent = x.Die1.PXNoOfComponent,
                    HotRunner = x.Die1.HotRunner,
                    GateType = x.Die1.GateType,
                    MCsize = x.Die1.MCsize,
                    CavQuantity = x.Die1.CavQuantity,
                    DieMakeLocation = x.Die1.DieMakeLocation,
                    DieMaker = x.Die1.DieMaker,
                    Family_Die_With = x.Die1.Family_Die_With,
                    Common_Part_With = x.Die1.Common_Part_With,
                    SpecialSpec = x.Die1.SpecialSpec,
                    DSUM_Idea = x.Die1.DSUM_Idea,
                    Q = x.Die1.NOofForQ,
                    C = x.Die1.NOofForC,
                    D = x.Die1.NOofForD,

                    MR_Request_Date = x.Die1.MR_Request_Date.HasValue ? x.Die1.MR_Request_Date.Value.ToString("MM/dd/yyyy") : null,
                    MR_App_Date = x.Die1.MR_App_Date.HasValue ? x.Die1.MR_App_Date.Value.ToString("MM/dd/yyyy") : null,
                    PO_Issue_Date = x.Die1.PO_Issue_Date.HasValue ? x.Die1.PO_Issue_Date.Value.ToString("MM/dd/yyyy") : null,
                    PO_App_Date = x.Die1.PODate.HasValue ? x.Die1.PODate.Value.ToString("MM/dd/yyyy") : null,
                    Design_Check_Plan = x.Die1.Design_Check_Plan,
                    Design_Check_Actual = x.Die1.Design_Check_Actual,
                    Design_Check_Result = x.Die1.Design_Check_Result,
                    NoOfPoit_Not_FL_DMF = x.Die1.NoOfPoit_Not_FL_DMF,
                    NoOfPoint_Not_FL_Spec = x.Die1.NoOfPoint_Not_FL_Spec,
                    NoOfPoint_Kaizen = x.Die1.NoOfPoint_Kaizen,
                    JIG_Using = x.Die1.JIG_Using == null ? null : x.Die1.JIG_Using == true ? "Yes" : "No",
                    JIG_No = x.Die1.JIG_No,
                    JIG_Check_Plan = x.Die1.JIG_Check_Plan,
                    JIG_Check_Result = x.Die1.JIG_Check_Result,
                    JIG_ETA_Supplier = x.Die1.JIG_ETA_Supplier,
                    JIG_Status = x.Die1.JIG_Status,
                    T0_Plan = x.Die1.T0_Plan.HasValue ? x.Die1.T0_Plan.Value.ToString("MM/dd/yyyy") : null,
                    T0_Actual = x.Die1.T0_Actual.HasValue ? x.Die1.T0_Actual.Value.ToString("MM/dd/yyyy") : null,
                    T0_Result = x.Die1.T0_Result,
                    T0_Solve_Method = x.Die1.T0_Solve_Method,
                    T0_Solve_Result = x.Die1.T0_Solve_Result,
                    Texture_Meeting_Date = x.Die1.Texture_Meeting_Date,
                    Texture_Go_Date = x.Die1.Texture_Go_Date,
                    S0_Plan = x.Die1.S0_Plan.HasValue ? x.Die1.S0_Plan.Value.ToString("MM/dd/yyyy") : null,
                    S0_Result = x.Die1.S0_Result,
                    S0_Solve_Method = x.Die1.S0_Solve_Method,
                    S0_solve_Result = x.Die1.S0_solve_Result,
                    Texture_App_Date = x.Die1.Texture_App_Date,
                    Texture_Internal_App_Result = x.Die1.Texture_Internal_App_Result,
                    Texture_JP_HP_App_Result = x.Die1.Texture_JP_HP_App_Result,
                    Texture_Note = x.Die1.Texture_Note,
                    PreKK_Plan = x.Die1.PreKK_Plan,
                    PreKK_Actual = x.Die1.PreKK_Actual,
                    PreKK_Result = x.Die1.PreKK_Result,
                    FA_Sub_Time = (x.Die1.FA_Sub_Time == null || x.Die1.FA_Sub_Time == 0) ? 0 : x.Die1.FA_Sub_Time - 1,
                    FA_Plan = x.Die1.FA_Plan.HasValue ? x.Die1.FA_Plan.Value.ToString("MM/dd/yyyy") : null,
                    FA_Result = x.Die1.FA_Result,
                    FA_Result_Date = x.Die1.FA_Result_Date.HasValue ? x.Die1.FA_Result_Date.Value.ToString("MM/dd/yyyy") : null,
                    FA_Problem = x.Die1.FA_Problem,
                    FA_Action_Improve = x.Die1.FA_Action_Improve,
                    MT1_Date = x.Die1.MT1_Date,
                    MT1_Gather_Date = x.Die1.MT1_Gather_Date,
                    MT1_Problem = x.Die1.MT1_Problem,
                    MT1_Remark = x.Die1.MT1_Remark,
                    MTF_Date = x.Die1.MTF_Date,
                    MTF_Gather_Date = x.Die1.MTF_Gather_Date,
                    MTF_Problem = x.Die1.MTF_Problem,
                    MTF_Remark = x.Die1.MTF_Remark,
                    First_Lot_Date = x.Die1.First_Lot_Date,
                    DieCost_USD = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_USD.HasValue ? String.Format(" {0:N}", x.Die1.DieCost_USD) : "***" : "***",
                    DieCost_JPY = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_JPY.HasValue ? String.Format(" {0:N}", x.Die1.DieCost_JPY) : "***" : "***",
                    DieCost_VND = commonFunction.isCanSeePrice(userID) == true ? x.Die1.DieCost_VND.HasValue ? String.Format(" {0:N}", x.Die1.DieCost_VND) : "***" : "***",
                    Warranty = x.Die1.DieWarranty_Short.HasValue ? String.Format(" {0:N0}", x.Die1.DieWarranty_Short) : "",
                    WarrantyShotAsDSUM = x.Die1.WarrantyShotAsDSUM.HasValue ? String.Format(" {0:N0}", x.Die1.WarrantyShotAsDSUM) : "",
                    RecordDate = x.Die1.RecordDate.HasValue ? x.Die1.RecordDate.Value.ToString("yyyy-MM-dd") : "-",
                    ActualShot = x.Die1.Short.HasValue ? String.Format(" {0:N0}", x.Die1.Short) : "",
                    Die_Status = db.DieStatusCategories.Find(x.Die1.DieStatusID)?.Type,

                    InventoryStatus = x.Die1.InventoryStatus,
                    RemarkDieStatusUsing = x.Die1.RemarkDieStatusUsing,
                    StartUse = x.Die1.StartOfUse.HasValue ? x.Die1.StartOfUse.Value.ToString("yyyy/MM/dd") : "-",
                    StopDate = x.Die1.StopDate.HasValue ? x.Die1.StopDate.Value.ToString("yyyy/MM/dd") : "-",
                    ReasonProposal = x.Die1.ReasonProposal,
                    SupplierProposal = x.Die1.SupplierProposal,
                    // TransferHis = db.DieLendingRequests.AsNoTracking().Where(trs => trs.DieNo == x.DieNo && trs.Active != false).ToList(),

                    Latest_Update_By = x.Die1.EditBy,
                    Latest_Update_Date = x.Die1.EditDate.HasValue ? x.Die1.EditDate.Value.ToString("MM/dd/yyyy") : null,
                    His_Update = x.Die1.His_Update,
                    Latest_Pending_Status_Changed = x.Die1.Latest_Pending_Status_Changed.HasValue ? x.Die1.Latest_Pending_Status_Changed.Value.ToString("MM/dd/yyyy") : null,
                    isCancel = x.Die1.isCancel,
                    isClosed = x.Die1.isClosed,
                    isActive = x.Die1.Active,
                    Attachment = new
                    {
                        DFM = db.Attachments.AsNoTracking().Where(c => c.DieID == x.DieID).Select(y => new
                        {
                            Clasify = y.Clasify,
                            FileName = y.FileName,
                            CreateDate = y.CreateDate,
                            CreateBy = y.CreateBy,
                            ReviseDate = y.ReviseDate,
                            ReviseBy = y.ReviseBy,
                            AttachID = y.AttachID,
                            Lastest = y.ReviseDate != null ? y.ReviseDate : y.CreateDate

                        }).OrderByDescending(y => y.Lastest).FirstOrDefault(),

                        Other = db.Attachments.AsNoTracking().Where(c => c.DieID == x.DieID && c.DFMID == null).ToList(),

                        TPI = db.Troubles.Where(c => c.DieID == x.DieID && c.Active != false && c.FinalStatusID != 11).Select(z => new
                        {
                            Clasify = z.TroubleName,
                            FileName = z.Report,
                            CreateDate = z.SubmitDate,
                            CreateBy = z.SubmitBy,
                            ReviseDate = "",
                            ReviseBy = "",
                            AttachID = z.TroubleID,
                            Lastest = z.PAECommentDate != null ? z.PAECommentDate : z.DMTAppDate
                        }).ToList() // 11: reject
                    }
                });
            }

            return outPut;
        }


        public JsonResult countForEachModel()
        {
            List<ModelList> listModel = db.ModelLists.ToList();
            List<object> datacounted = new List<object>();
            var allDie = db.Die1.Where(x => x.Active != false && x.isCancel != true && x.DieClassify.Contains("MT"));

            foreach (var model in listModel)
            {
                int c = allDie.Where(x => x.ModelID == model.ModelID).Count();
                if (c > 0)
                {
                    datacounted.Add(new
                    {
                        modelName = model.ModelName,
                        modelID = model.ModelID,
                        Qty = c
                    });
                }
            }

            return Json(datacounted, JsonRequestBehavior.AllowGet);
        }














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
