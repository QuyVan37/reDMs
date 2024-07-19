using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using Aspose.Slides;
using Aspose.Cells;
using Aspose.Pdf;
using Avalonia.Controls;
using DMS03.Models;

using iTextSharp.text;
using OfficeOpenXml;
using Spire.Xls;

using Spire.Presentation;
using static iTextSharp.text.pdf.AcroFields;
using Newtonsoft.Json;
using System.Web.UI.WebControls;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using System.Web.Routing;
using OfficeOpenXml.ConditionalFormatting;
using System.Runtime.InteropServices.ComTypes;

using Aspose.Slides.Export.Web;
using iTextSharp.text.pdf;
using System.Text;
using System.Runtime.Remoting.Messaging;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Excel;
using Spire.Pdf.Graphics;

using Spire.Pdf;
using System.Runtime.ConstrainedExecution;
using System.Web.Configuration;
using ceTe;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;
using Microsoft.Office.Core;
using System.Threading.Tasks;
using ceTe.DynamicPDF.Merger;

using System.Windows.Controls;
using Rectangle = iTextSharp.text.Rectangle;

using Aspose.Slides.Export;
using System.Windows;

using pdftron;
using pdftron.Common;
using pdftron.SDF;
using pdftron.PDF;
using Aspose.Pdf.Operators;
using static pdftron.PDF.Convert;
using static iTextSharp.awt.geom.Point2D;
using Org.BouncyCastle.Asn1;
using Spire.Doc;
using System.Web.Services.Description;
using System.IO.Packaging;
using Microsoft.Ajax.Utilities;

namespace DMS03.Controllers
{
    public class DSUMsController : Controller
    {
        private DMSEntities db = new DMSEntities();
        public SendEmailController mailJob = new SendEmailController();
        public CommonFunctionController commonFunction = new CommonFunctionController();
        StoreProcudure storeProcudure = new StoreProcudure();

        public DateTime applyDate = new DateTime(2023, 12, 13);
        // GET: DSUMs


        public ActionResult Index()
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }

            //var list = db.DSUMs.ToList();
            ViewBag.SupplierID = new SelectList(db.Suppliers, "SupplierID", "SupplierName");
            ViewBag.ModelID = new SelectList(db.ModelLists, "ModelID", "ModelName");
            ViewBag.ProcessCodeID = new SelectList(db.ProcessCodeCalogories, "ProcessCodeID", "Type");
            //ViewBag.WPAEPICCheck = db.DSUMs.Where(x => x.DSUMStatusID == 1).Count();
            ////ViewBag.WPAEG4UPCheck = db.DSUMs.Where(x => x.DSUMStatusID == 2).Count();
            //ViewBag.WPE1PICCheck = db.DSUMs.Where(x => x.DSUMStatusID == 2).Count();
            ////ViewBag.WPE1G4UPCheck = db.DSUMs.Where(x => x.DSUMStatusID == 4).Count();
            //ViewBag.WPE1Approve = db.DSUMs.Where(x => x.DSUMStatusID == 3).Count();
            //ViewBag.WPAEApprove = db.DSUMs.Where(x => x.DSUMStatusID == 4).Count();
            //ViewBag.Rejected = db.DSUMs.Where(x => x.DSUMStatusID == 6).Count();
            //ViewBag.Cancelled = db.DSUMs.Where(x => x.DSUMStatusID == 7).Count();
            //ViewBag.Finhished = db.DSUMs.Where(x => x.DSUMStatusID == 5).Count();
            //var dSUMs = db.DSUMs.Include(d => d.DSUMStatusCategory);
            return View();
        }


        public JsonResult DownloadListDSUMAll()
        {

            var output = storeProcudure.getListDSUMAll("");

            return Json(output, JsonRequestBehavior.AllowGet);
        }



        public JsonResult getData(string waitfor, string search, string[] supplierID, string[] modelID, string[] procesCodeID, string from, string to, string export, int? page)
        {

            int pageIndex = page ?? 1;
            int pageSize = 50;

            modelID = modelID == null ? new string[0] { } : modelID.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            supplierID = supplierID == null ? new string[0] { } : supplierID.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            procesCodeID = procesCodeID == null ? new string[0] { } : procesCodeID.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            List<DSUM> list = new List<DSUM>();

            if (!String.IsNullOrEmpty(waitfor))
            {
                if (waitfor.Contains("no-route"))
                {
                    list = db.DSUMs.Where(x => (x.Route == null || x.Route.Trim() == String.Empty) && x.DSUMStatusID != 11).ToList();
                }
                else
                {
                    list = db.DSUMs.Where(x => x.DSUMStatusCategory.Status.Contains(waitfor)).ToList();
                }

                goto next;
            }

            if (!String.IsNullOrEmpty(search))
            {
                list = db.DSUMs.Where(x => x.PartNo.Contains(search)).ToList();
                if (list == null)
                {
                    list = db.DSUMs.Where(x => x.DSUMNo.Contains(search)).ToList();
                }
                goto next;
            }

            if (modelID.Length > 0)
            {
                List<DSUM> searchResult = new List<DSUM>();
                foreach (var id in modelID)
                {
                    int intID = int.Parse(id);
                    var res = db.DSUMs.Where(x => x.Die1.ModelID == intID).ToList();
                    searchResult.AddRange(res);
                }
                list = searchResult;
                if (list.Count() == 0) goto next;

            }
            if (supplierID.Length > 0)
            {
                List<DSUM> searchResult = new List<DSUM>();
                foreach (var id in supplierID)
                {
                    if (list.Count() == 0)
                    {
                        int intID = int.Parse(id);
                        var res = db.DSUMs.Where(x => x.Die1.SupplierID == intID).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        int intID = int.Parse(id);
                        var res = list.Where(x => x.Die1.SupplierID == intID).ToList();
                        searchResult.AddRange(res);
                    }
                }
                list = searchResult;
                if (list.Count() == 0) goto next;

            }

            if (procesCodeID.Length > 0)
            {
                List<DSUM> searchResult = new List<DSUM>();
                foreach (var id in procesCodeID)
                {
                    if (list.Count() == 0)
                    {
                        int intID = int.Parse(id);
                        var res = db.DSUMs.Where(x => x.Die1.ProcessCodeID == intID).ToList();
                        searchResult.AddRange(res);
                    }
                    else
                    {
                        int intID = int.Parse(id);
                        var res = list.Where(x => x.Die1.ProcessCodeID == intID).ToList();
                        searchResult.AddRange(res);
                    }
                }
                list = searchResult;
                if (list.Count() == 0) goto next;

            }


            if (!String.IsNullOrWhiteSpace(from))
            {
                List<DSUM> searchResult = new List<DSUM>();
                DateTime fromDate = DateTime.Parse(from);
                if (list.Count() == 0)
                {
                    var res = db.DSUMs.Where(x => x.SubmitDate >= fromDate).ToList();
                    searchResult.AddRange(res);
                }
                else
                {

                    var res = list.Where(x => x.SubmitDate >= fromDate).ToList();
                    searchResult.AddRange(res);
                }

                list = searchResult;
                if (list.Count() == 0) goto next;
            }
            if (!String.IsNullOrWhiteSpace(to))
            {
                List<DSUM> searchResult = new List<DSUM>();
                DateTime todate = DateTime.Parse(to);
                if (list.Count() == 0)
                {
                    var res = db.DSUMs.Where(x => x.SubmitDate <= todate).ToList();
                    searchResult.AddRange(res);
                }
                else
                {

                    var res = list.Where(x => x.SubmitDate <= todate).ToList();
                    searchResult.AddRange(res);
                }

                list = searchResult;
                if (list.Count() == 0) goto next;
            }


            if (!String.IsNullOrEmpty(export))
            {
                //goto export
            }

        next:
            var result = list.Where(x => x.Active != false).OrderByDescending(x => x.SubmitDate).Page(pageIndex, pageSize).Select(x => new
            {
                DFMID = x.DFMID,
                PartNo = x.PartNo,
                DieNo = x.DieNo,
                Status = x.DSUMStatusCategory.Status,
                DSUMNo = x.DSUMNo,
                Warning = x.DSUMStatusID != 9 && x.DSUMStatusID != 10 && x.DSUMStatusID == 11 ? "Already Keeping " + (DateTime.Now.Subtract((DateTime)x.LastestReviseDate).Days - commonFunction.countHoliday((DateTime)x.LastestReviseDate, DateTime.Now)).ToString() + " days" : "-",
                DFMFinished = x.DSUMStatusID == 9 ? (db.Attachments.Where(y => y.DFMID == x.DFMID && (y.Clasify == "DFM PAE-APPROVED" || y.Clasify == "DFM PAE-G6UP-Approve" || y.Clasify == "DFM_Finished")).FirstOrDefault() == null ? "" : db.Attachments.Where(y => y.DFMID == x.DFMID && (y.Clasify == "DFM PAE-APPROVED" || y.Clasify == "DFM PAE-G6UP-Approve" || y.Clasify == "DFM_Finished")).FirstOrDefault()?.FileName) : "",
                Supplier = x.Die1.Supplier.SupplierName,
                Model = x.Die1.ModelList.ModelName,
                Submitor = x.SubmitBy,
                SubmitDate = x.SubmitDate.HasValue ? x.SubmitDate.Value.ToString("yyyy/MM/dd") : "",
                ApproveDate = x.PAEApproveDate.HasValue ? x.PAEApproveDate.Value.ToString("yyyy/MM/dd") : "",
                Remark = x.Remark,
                POdate = db.Die1.Where(y => y.PartNoOriginal == x.PartNo && y.Die_Code == x.DieNo && y.Active != false && y.isCancel != true).FirstOrDefault()?.PODate
            }).ToList();

            double totalPage = decimal.ToDouble(list.Count()) / decimal.ToDouble(pageSize);
            var output = new
            {
                test = totalPage,
                page = pageIndex,
                totalPage = Math.Ceiling(totalPage),
                data = result,
            };

            return Json(output, JsonRequestBehavior.AllowGet);
        }

        //public class SupplierAndModelAndProcess
        //{
        //    public string supplier { set; get; }
        //    public int? supplierID { set; get; }

        //    public string model { set; get; }
        //    public int? modelID { set; get; }
        //    public string processCode { set; get; }
        //    public int? processCodeID { set; get; }
        //}
        //public SupplierAndModelAndProcess getSupplierAndModelAndProcessCode(string partNo, string dieNo)
        //{

        //    var supplier = "Plz Update [New Die Launching]";
        //    var model = "Plz Update [New Die Launching]";
        //    var item = db.Die1.Where(x => x.PartNoOriginal == partNo && x.Die_Code == dieNo && x.Active != false && x.isCancel != true).FirstOrDefault();
        //    int supplierID = 0;
        //    int modelID = 0;
        //    var processCode = "";
        //    int processCodeID = 0;

        //    if (item != null)
        //    {
        //        supplier = item.SupplierID != null ? db.Suppliers.Find(item.SupplierID).SupplierName : supplier;
        //        try
        //        {
        //            supplierID = (int)item.SupplierID;
        //        }
        //        catch
        //        {
        //            supplierID = 0;
        //        }
        //        model = item.ModelID != null ? db.ModelLists.Find(item.ModelID).ModelName : model;
        //        try
        //        {
        //            modelID = (int)item.ModelID;
        //        }
        //        catch
        //        {
        //            modelID = 0;
        //        }

        //        processCode = item.ProcessCodeID != null ? db.ProcessCodeCalogories.Find(item.ProcessCodeID).Type : processCode;
        //        try
        //        {
        //            processCodeID = (int)item.ProcessCodeID;
        //        }
        //        catch
        //        {
        //            processCodeID = 0;
        //        }
        //    }

        //    SupplierAndModelAndProcess output = new SupplierAndModelAndProcess()
        //    {
        //        supplier = supplier,
        //        model = model,
        //        supplierID = supplierID,
        //        modelID = modelID,
        //        processCode = processCode,
        //        processCodeID = processCodeID
        //    };
        //    return output;
        //}

        public JsonResult summarize()
        {
            int wMeeting = db.DSUMs.Where(x => x.DSUMStatusID == 1 && x.Active != false).Count();
            int wPAEG4 = db.DSUMs.Where(x => x.DSUMStatusID == 2 && x.Active != false).Count();
            int wDMTG6 = db.DSUMs.Where(x => x.DSUMStatusID == 3 && x.Active != false).Count();
            int wPE1G4 = db.DSUMs.Where(x => x.DSUMStatusID == 4 && x.Active != false).Count();
            int wJPPAE = db.DSUMs.Where(x => x.DSUMStatusID == 5 && x.Active != false).Count();
            int wJPPE1 = db.DSUMs.Where(x => x.DSUMStatusID == 6 && x.Active != false).Count();
            int wPE1App = db.DSUMs.Where(x => x.DSUMStatusID == 7 && x.Active != false).Count();
            int wPAEApp = db.DSUMs.Where(x => x.DSUMStatusID == 8 && x.Active != false).Count();
            int finished = db.DSUMs.Where(x => x.DSUMStatusID == 9 && x.Active != false).Count();
            int Rejected = db.DSUMs.Where(x => x.DSUMStatusID == 10 && x.Active != false).Count();
            int cancelled = db.DSUMs.Where(x => x.DSUMStatusID == 11 && x.Active != false).Count();
            int wPE1PIC = db.DSUMs.Where(x => x.DSUMStatusID == 12 && x.Active != false).Count();
            int wECN = db.DSUMs.Where(x => x.DSUMStatusID == 13 && x.Active != false).Count();
            int noRoute = db.DSUMs.Where(x => (x.Route == null || x.Route.Trim() == String.Empty) && x.Active != false && x.DSUMStatusID != 11).Count();
            var output = new
            {
                noRoute = noRoute,
                wMeeting = wMeeting,
                wPAEG4 = wPAEG4,
                wDMTG6 = wDMTG6,
                wPE1G4 = wPE1G4,
                wJPPAE = wJPPAE,
                wJPPE1 = wJPPE1,
                wPE1App = wPE1App,
                wPAEApp = wPAEApp,
                finished = finished,
                Rejected = Rejected,
                cancelled = cancelled,
                wPE1PIC = wPE1PIC,
                wECN = wECN
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }


        public ContentResult getDetail(int id)
        {
            db.Configuration.ProxyCreationEnabled = false;
            var record = db.DSUMs.Find(id);
            var attachment = db.Attachments.Where(x => x.DFMID == record.DFMID).Select(y => new
            {
                Clasify = y.Clasify,
                FileName = y.FileName,
                CreateDate = y.CreateDate,
                CreateBy = y.CreateBy,
                ReviseDate = y.ReviseDate,
                ReviseBy = y.ReviseBy,
                AttachID = y.AttachID,
                Lastest = y.ReviseDate != null ? y.ReviseDate : y.CreateDate

            }).OrderByDescending(y => y.Lastest).ToList();
            var status = db.DSUMStatusCategories.AsNoTracking().Where(x => x.DSUMStatusID == record.DSUMStatusID).FirstOrDefault();
            var output = new
            {
                record = new
                {
                    DFMID = record.DFMID,
                    DieID = record.DieID,
                    DSUMStatusID = record.DSUMStatusID,
                    Route = record.Route,
                    PartNo = record.PartNo,
                    DieNo = record.DieNo,
                    Remark = record.Remark,
                    Role = getRole(id)
                },
                attachment = attachment,
                status = status,
                routeAndPIC = getRouteAndPIC(record.DFMID)
            };

            // return Json(output, JsonRequestBehavior.AllowGet);

            var list = JsonConvert.SerializeObject(output,
                     Formatting.None,
                    new JsonSerializerSettings()
                    {
                        ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore
                    });

            return Content(list, "application/json");

        }

        public JsonResult newDFM(string partNo, string dieNo, HttpPostedFileBase fileDFM, HttpPostedFileBase coverpage)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["DSUM_Role"].ToString();
            var name = Session["name"].ToString();
            if (role != "Check" && role != "Approve" && role != "Issue")
            {
                return Json(new { status = false, msg = "No permision!" }, JsonRequestBehavior.AllowGet);
            }

            var today = DateTime.Now;
            var msg = "";
            bool status = false;
            if (!String.IsNullOrWhiteSpace(partNo) && !String.IsNullOrWhiteSpace(dieNo) && fileDFM != null && coverpage != null)
            {
                //1. Xử lí input
                var PartNo = partNo.ToUpper().Trim();
                verify verifyResult = verifyFormCoverPage(coverpage, PartNo, "", true);
                if (verifyResult.isPass == false)
                {
                    msg = verifyResult.msg;
                    status = false;
                    goto next;
                }

                var DieNo = dieNo.ToUpper().Trim();
                var fileExtCoverpage = Path.GetExtension(coverpage.FileName);
                var fileExtDFMContent = Path.GetExtension(fileDFM.FileName);
                if (!fileExtCoverpage.ToLower().Contains(".xls")) // file excel
                {
                    return Json(new { status = false, msg = "Coverpage must excel" }, JsonRequestBehavior.AllowGet);
                }
                if (!fileExtDFMContent.ToLower().Contains(".ppt") && !fileExtDFMContent.ToLower().Contains(".pdf")) // Không phải ppt và pdf
                {
                    return Json(new { status = false, msg = "DFM content must pptx" }, JsonRequestBehavior.AllowGet);
                }

                //2. Check đã tồn tại chưa (ngoại trừ con bị xóa và Cancel
                var isExistDFM = db.DSUMs.Where(x => x.PartNo == PartNo && x.DieNo == DieNo && x.Active != false && x.DSUMStatusID != 10 && x.DSUMStatusID != 11).FirstOrDefault(); // 10/11: Reject/Cancel
                if (isExistDFM != null) // Đã tồn tại DFM này rồi
                {
                    msg = "This DFM already exist! It was submited by " + isExistDFM.SubmitBy + "/ " + isExistDFM.SubmitDate.Value.ToShortDateString();
                    status = false;
                    goto next;
                }
                //3. Check đã có bên New Die Launching chưa?
                var die = db.Die1.Where(x => x.PartNoOriginal == PartNo && x.Die_Code == DieNo && x.Active != false && x.isCancel != true).FirstOrDefault();
                if (die == null) // Chưa có trong danh sách New Die Launching => Y/c input vào List New Die Lauching first
                {

                    mailJob.sendEmailToPAEAnnouceUploadListNewDie(PartNo, DieNo);
                    msg = "This die not exist in [Die Information]! Plz re-Check PartNo, DieNo is correct or not?. If it correct plz go to [Die Information] and click [Add New] ";
                    status = false;
                    goto next;
                }
                else
                {
                    if (die.DieClassify == "MP")
                    {
                        var otherDie = db.Die1.Where(x => x.PartNoOriginal == PartNo && x.Active != false && x.isCancel != true && x.DieClassify != "MP").FirstOrDefault();
                        if (otherDie != null)
                        {
                            die = otherDie;
                        }
                    }

                }


                //4 Lưu File vật lí
                //4.1 Lưu cover page
                var fileNameCoverPage = "DFM_Sub" + "_" + die.DieNo + today.ToString("_yyyy-MM-dd-HHmmss") + fileExtCoverpage;
                var pathCover = Path.Combine(Server.MapPath("~/File/Attachment/"), fileNameCoverPage);


                coverpage.SaveAs(pathCover);

                //5.2 Lưu DFM content
                var fileExtDFM = Path.GetExtension(fileDFM.FileName);
                var fileNameDFM = "DFM_Content_Sub" + "_" + die.DieNo + today.ToString("_yyyy-MM-dd-HHmmss") + fileExtDFM;
                var pathDFM = Path.Combine(Server.MapPath("~/File/Attachment/"), fileNameDFM);
                fileDFM.SaveAs(pathDFM);
                // 5.3. insert DFMfile to cover page
                string srcImage = Path.Combine(Server.MapPath("~/File/Attachment/"), "DFMContentDefaulImage.png");

                Bitmap bitmap = new Bitmap(srcImage);
                byte[] imageData = null;
                byte[] fileEmbed = null;

                // get file to make icon
                using (FileStream fs = new FileStream(srcImage, FileMode.Open, FileAccess.Read))
                {
                    // Create a byte array of file stream length
                    imageData = System.IO.File.ReadAllBytes(srcImage);
                    //Read block of bytes from stream into the byte array
                    fs.Read(imageData, 0, System.Convert.ToInt32(fs.Length));
                    //Close the File Stream
                    fs.Close();
                }

                // get file DFM content
                using (FileStream fs = new FileStream(pathDFM, FileMode.Open, FileAccess.Read))
                {
                    // Create a byte array of file stream length
                    fileEmbed = System.IO.File.ReadAllBytes(pathDFM);
                    //Read block of bytes from stream into the byte array
                    fs.Read(fileEmbed, 0, System.Convert.ToInt32(fs.Length));
                    //Close the File Stream
                    fs.Close();
                }


                using (Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(pathCover))
                {
                    Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                    if (worksheet == null)
                    {
                        worksheet = workbook.Worksheets["Att-01_ver7__Mo_(Update size)"];
                    }
                    if (worksheet == null)
                    {
                        worksheet = workbook.Worksheets["Att-02_ver8__PX_(Official)"];
                    }
                    if (worksheet == null)
                    {
                        worksheet = workbook.Worksheets["Att-02_ver7__PX_(Official)"];
                    }
                    if (worksheet == null)
                    {
                        worksheet = workbook.Worksheets["MO"];
                    }
                    //string A1Value = worksheet.Cells["A1"].Value.ToString();
                    //if (A1Value.ToUpper().Contains("MO"))
                    //{
                    worksheet.Shapes.AddOleObject(36, 0, 10, 0, 150, 300, imageData);
                    //}
                    //else
                    //{
                    //    worksheet.Shapes.AddOleObject(38, 0, 18, 0, 100, 200, imageData);
                    //}


                    // Set embedded ole object data.
                    worksheet.OleObjects[0].ObjectData = fileEmbed;
                    worksheet.OleObjects[0].DisplayAsIcon = true;
                    workbook.Save(pathCover);

                };
                deleteWorksheetLisence(pathCover);


                //5. Mọi thứ đã OK => Luu vào database
                //5. Luu DSUM infor
                DSUM newDFM = new DSUM()
                {
                    DSUMStatusID = 1, // W-PAE-PIC-Check
                    DieID = die.DieID,
                    PartNo = PartNo,
                    DieNo = die.Die_Code,
                    DSUMNo = genarateDSUMNo("", true, false),
                    SubmitBy = Session["Name"].ToString(),
                    SubmitDate = today,
                    SubmitorEmail = Session["Mail"].ToString(),
                    CreateDate = today,
                    LastestReviseDate = today
                };

                // update Cover page => lastest version
                bool isSuccessSign = false;
                string controlLatestVersion = "By " + name + "_" + today.ToString("yyyy/MM/dd HH:mm:ss tt");
                if (today > applyDate)
                {
                    isSuccessSign = getSignatureAndSaveCoverPage(pathCover, 1, name, false, true, "", controlLatestVersion, false, null, die.Die_Code);
                    if (isSuccessSign)
                    {
                        newDFM.ControlLastestDFM = controlLatestVersion;
                    }
                    else
                    {
                        goto next;
                    }
                }

                bool isSuccessCover = handelDSMtoPDF(fileNameCoverPage);
                if (isSuccessCover == false)
                {
                    status = false;
                    msg = "Coverpage is not insert(Embed) DFM Content";
                    goto next;
                }
                db.DSUMs.Add(newDFM);
                db.SaveChanges();

                Attachment newAtt = new Attachment()
                {
                    DieID = die.DieID,
                    DFMID = newDFM.DFMID,
                    FileName = fileNameCoverPage,
                    Clasify = "DFM_SUPPLIER_SUBMIT",
                    CreateBy = Session["Name"].ToString(),
                    CreateDate = today,
                };
                db.Attachments.Add(newAtt);

                // 3.3 Update New Die Launching
                die.DFM_Sub_Date = today.ToString("MM/dd/yyyy");
                die.AttDFMMaker = fileNameCoverPage;
                die.Genaral_Information = today.ToString("MM/dd/yyyy") + ": DFM Submited " + System.Environment.NewLine + die.Genaral_Information;
                db.Entry(die).State = EntityState.Modified;
                db.SaveChanges();
                deleteFileStoreInTempFolder(pathDFM);
                status = true;
                msg = "Success";
                //5.send email to PAE
                mailJob.sendEmailCheckDFM(newDFM, false, "");
            }
            else
            {
                msg = "Plz input enough infor & 12 letter Part No & 3 letter Die No";
            }

        next:
            var outPut = new
            {
                status = status,
                msg = msg
            };


            return Json(outPut, JsonRequestBehavior.AllowGet);
        }


        public ActionResult RPAissueNewDFM()
        {
            Session["Name"] = "RPA";
            Session["Mail"] = "quy.nguyen324@mail.canon";
            Session["Dept"] = "PAE";
            Session["DSUM_Role"] = "Issue";

            return View();
        }




        public string genarateDSUMNo(string currentDTFNo, bool isNew, bool isRevise)
        {
            var output = "";
            var today = DateTime.Now;
            var totalDSUMinThisYear = db.DSUMs.Where(x => x.SubmitDate.Value.Year == today.Year).Count() + 1;
            if (isNew)
            {
                output = "DFM" + today.ToString("yyMMdd-") + totalDSUMinThisYear + "-00";
            }
            if (isRevise)
            {
                var upver = currentDTFNo.Substring(currentDTFNo.Length - 2, 2); // 00
                int upverInt = System.Convert.ToInt16(upver) + 1;
                string upverStr = System.Convert.ToString(upverInt);
                if (upverStr.Length == 1)
                {
                    upverStr = "0" + upverStr;
                }
                var mainNo = currentDTFNo.Remove(currentDTFNo.Length - 2, 2);
                output = mainNo + upverStr;
            }
            return output;
        }

        public JsonResult getRouteAndPIC(int? DFMID)
        {
            db.Configuration.ProxyCreationEnabled = false;
            List<object> output = new List<object>();
            if (DFMID == null)
            {
                var allStatusRoute = db.DSUMStatusCategories.Where(x => x.OrderPriority != null && x.DSUMStatusID != 9).OrderBy(x => x.OrderPriority).ToList();
                foreach (var item in allStatusRoute)
                {
                    output.Add(new
                    {
                        route = new
                        {
                            RouteID = item.DSUMStatusID,
                            RouteName = item.RouteName,

                        },
                        pics = db.Users.Where(x => x.Department.DeptName.Contains(item.DeptResponse) && x.DSUMRole.Contains(item.RoleResponse) && item.GradeResponse.Contains(x.Grade) && x.Active != false).Select(u => new
                        {
                            UsderID = u.UserID,
                            UserName = u.UserName
                        })
                    });
                }
            }
            else
            {
                var dsum = db.DSUMs.Find(DFMID);
                if (String.IsNullOrEmpty(dsum.Route))
                {
                    output.Add(new
                    {
                        route = new
                        {
                            RouteID = 0,
                            RouteName = "You did not select route yet!"
                        },
                        pics = new List<object>()
                    });
                }
                else
                {
                    string[] routeIDstring = dsum.Route.Split(',');
                    foreach (var idSring in routeIDstring)
                    {
                        List<object> listUser = new List<object>();
                        int id = int.Parse(idSring);
                        var item = db.DSUMStatusCategories.Find(id);
                        var pics = db.DSUM_PIC.Where(x => x.DFMID == dsum.DFMID && x.ProcessID == idSring).FirstOrDefault();
                        if (pics == null)
                        {
                            if (id != 9 && id != 13)
                            {
                                listUser.Add(new
                                {
                                    UserID = 0,
                                    UserName = "PAE(CVN/JP)"
                                });
                                listUser.Add(new
                                {
                                    UserID = 0,
                                    UserName = "PE1(CVN/JP)"
                                });
                                listUser.Add(new
                                {
                                    UserID = 0,
                                    UserName = "DMT(ifAny)"
                                });
                            }
                            else
                            {
                                listUser = new List<object>();
                            }


                        }
                        else
                        {
                            string[] listPICUserIDstring = pics.UserListPIC.Split(',');
                            foreach (var userID in listPICUserIDstring)
                            {
                                if (!String.IsNullOrWhiteSpace(userID))
                                {
                                    var user = db.Users.Find(int.Parse(userID));
                                    listUser.Add(new
                                    {
                                        UserID = user.UserID,
                                        UserName = user.UserName
                                    });
                                }

                            }
                        }

                        output.Add(new
                        {
                            route = new
                            {
                                RouteID = item.DSUMStatusID,
                                RouteName = item.RouteName,
                            },
                            pics = listUser
                        });
                    }
                }

                //output.Add(new
                //{
                //    route = new
                //    {
                //        RouteID = 9,
                //        RouteName = "Finished",
                //    },
                //    pics = new List<object>()
                //});

            }


            return Json(output, JsonRequestBehavior.AllowGet);
        }



        public class routePost
        {
            public routeConfig route { set; get; }
            public pic[] pics { set; get; }

        }

        public class pic
        {
            public int UserID { set; get; }
            public string UserName { set; get; }
        }
        public class routeConfig
        {
            public int RouteID { set; get; }
            public string RouteName { set; get; }
        }

        public JsonResult assignRoute(string[] ids, routePost[] route)
        {
            var dept = Session["Dept"].ToString();
            var grade = Session["Grade"].ToString();
            var role = Session["DSUM_Role"].ToString();
            var today = DateTime.Now;
            var name = Session["Name"].ToString();
            var status = false;
            string[] greadAllow = { "G4", "G5", "G6", "M1", "AGM", "GM" };
            if (route.Length < 2)
            {
                return Json(new { status = false, msg = "Please assign route" }, JsonRequestBehavior.AllowGet);
            }

            if (!greadAllow.Contains(grade))
            {
                return Json(new { status = false, msg = "Only G4 Up of PAE can assign route!" }, JsonRequestBehavior.AllowGet);
            }

            if (dept.Contains("PAE") && (role.Contains("Check") || role.Contains("Approve")) && greadAllow.Contains(grade))
            {
                // 1. Lay route va PIC

                foreach (var id in ids)
                {
                    int int_ID = int.Parse(id);
                    var dsum = db.DSUMs.Find(int_ID);

                    string DFMRoute = "";
                    for (var i = 1; i < route.Length; i++)
                    {
                        DFMRoute = DFMRoute + "," + route[i].route.RouteID;
                        string routePIC = "";
                        var pics = route[i].pics;
                        if (pics != null)
                        {
                            foreach (var pic in route[i].pics)
                            {
                                routePIC = routePIC + "," + pic?.UserID.ToString();
                            }
                        }

                        DSUM_PIC newPIC = new DSUM_PIC();
                        newPIC.DFMID = dsum.DFMID;
                        newPIC.ProcessID = route[i].route.RouteID.ToString();
                        newPIC.UserListPIC = routePIC.Length > 0 ? routePIC.Remove(0, 1) : routePIC;

                        // Kiem tra trong db DFMID va PRocessID đã có chưa
                        // Nếu chưa có => add new
                        // Nếu đã có thì modify
                        var existProcess = db.DSUM_PIC.Where(x => x.DFMID == newPIC.DFMID && x.ProcessID == newPIC.ProcessID).FirstOrDefault();
                        if (existProcess != null)
                        {
                            existProcess.UserListPIC = newPIC.UserListPIC;
                            db.Entry(existProcess).State = EntityState.Modified;
                            db.SaveChanges();
                        }
                        else
                        {
                            db.DSUM_PIC.Add(newPIC);
                            db.SaveChanges();
                        }
                    }

                    if (dsum.DieNo == "11A")
                    {
                        dsum.Route = "1" + DFMRoute + ",13,9";
                    }
                    else
                    {
                        dsum.Route = "1" + DFMRoute + ",9";
                    }
                    dsum.Remark = today.ToString("yyyy/MM/dd: ") + name + " : Assign DSUM Route!" + System.Environment.NewLine + dsum.Remark;
                    db.Entry(dsum).State = EntityState.Modified;
                    db.SaveChanges();
                    status = true;
                }
            }
            return Json(new { status = status, msg = "OK" }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult UpLoadDFM(int DFMID, HttpPostedFileBase fileDFM, string type, string remark)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["DSUM_Role"].ToString();
            var grade = Session["Grade"].ToString();
            if (role != "Check" && role != "Approve")
            {
                return Json(new { status = false, assignRoute = false, msg = "No permision!" }, JsonRequestBehavior.AllowGet);
            }


            var msg = "";
            var name = Session["Name"].ToString();
            var mail = Session["Mail"].ToString();
            var status = false;
            var today = DateTime.Now;
           
            var sender = mail;
           
            var item = "";
          



            if (DFMID != 0)
            {
                var DSUM = db.DSUMs.Find(DFMID);

                // Check role
                var statusConfig = db.DSUMStatusCategories.Find(DSUM.DSUMStatusID);
                //check form
                if (fileDFM == null && DSUM.DSUMStatusID != 1)
                {
                    if (statusConfig.DSUMStatusID == 1)
                    {

                    }

                    var attachment = db.Attachments.Where(x => x.DFMID == DSUM.DFMID).Select(y => new
                    {
                        Clasify = y.Clasify,
                        FileName = y.FileName,
                        CreateDate = y.CreateDate,
                        CreateBy = y.CreateBy,
                        ReviseDate = y.ReviseDate,
                        ReviseBy = y.ReviseBy,
                        AttachID = y.AttachID,
                        Lastest = y.ReviseDate != null ? y.ReviseDate : y.CreateDate

                    }).OrderByDescending(y => y.Lastest).FirstOrDefault();

                    {
                        string fullpathDFM = Path.Combine(Server.MapPath("~/File/Attachment/"), attachment.FileName);
                        byte[] bytes = System.IO.File.ReadAllBytes(fullpathDFM);
                        fileDFM = (HttpPostedFileBase)new MemoryPostedFile(bytes, fullpathDFM, "Copy" + attachment.FileName);

                    }


                }
                verify verifyResult = verifyFormCoverPage(fileDFM, DSUM.PartNo, DSUM.ControlLastestDFM, false);
                if (verifyResult.isPass == false)
                {
                    msg = verifyResult.msg;
                    status = false;
                    goto Next;
                }


                //***

                if ((dept.Contains(statusConfig.DeptResponse) || statusConfig.DeptResponse == "All") && statusConfig.GradeResponse.Contains(grade) && (statusConfig.RoleResponse == role || role == "Approve"))
                {
                    if (String.IsNullOrWhiteSpace(DSUM.Route))
                    {
                        return Json(new { status = false, assignRoute = true, msg = "Please assgin Route first!" }, JsonRequestBehavior.AllowGet);
                    }
                    item = DSUM.PartNo + "-" + DSUM.DieNo;
                    var die = db.Die1.Where(x => x.PartNoOriginal == DSUM.PartNo && x.Die_Code == DSUM.DieNo && x.Active != false && x.isCancel != true).FirstOrDefault();

                    string[] routeArray = DSUM.Route.Split(',');
                    int currentIndex = Array.IndexOf(routeArray, DSUM.DSUMStatusID.ToString());
                    int CurrentStatusID = (int)DSUM.DSUMStatusID;


                    if (fileDFM != null && type != null)
                    {


                        //if (currentIndex == routeArray.Length - 1)
                        //{
                        //    DSUM.DSUMStatusID = 9; // Finished

                        //}
                        //else
                        //{
                        //    DSUM.DSUMStatusID = int.Parse(routeArray[currentIndex + 1]); // New status
                        //}

                        //var newStatus = db.DSUMStatusCategories.Find(DSUM.DSUMStatusID);
                        // For mail

                        //reciever = newStatus.DeptResponse + newStatus.GradeResponse + "UP";
                        //mailmsg = "DFM neeed " + reciever + " check/Approve";
                        //deptRecive = newStatus.DeptResponse;


                        if (CurrentStatusID == 1) // W-meeting 
                        {
                            // PAE checked
                            DSUM.PAECheckby = name;
                            DSUM.PAECheckDate = today;
                            // PE1 Checked
                            DSUM.PE1CheckDate = today;
                            //die.DFM_PAE_Check_Date = today.ToString("yyyy/MM/dd");
                            //die.DFM_PE_Check_Date = today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": Done Meeting!" + System.Environment.NewLine + die.Genaral_Information;

                        }
                        if (CurrentStatusID == 2) // W-PAE-G4UP-Check
                        {
                            // PAE checked
                            DSUM.PAEG4UP_PIC = name + "/" + today.ToString("yyyy/MM/dd");
                            DSUM.PAECheckDate = today;
                            die.DFM_PAE_Check_Date = today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": PAE Checked!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": PAE Checked!" + System.Environment.NewLine + DSUM.Remark;
                        }
                        if (CurrentStatusID == 3) // W-DMT-G6UP-Check
                        {
                            DSUM.DMTG6UP_PIC = name + "/" + today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DMT Approved!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ":DMT Approved " + System.Environment.NewLine + DSUM.Remark;

                        }
                        if (CurrentStatusID == 4) // W-PE1- G4UP-Check
                        {
                            DSUM.PE1G4UP_PIC = name + "/" + today.ToString("yyyy/MM/dd");
                            die.DFM_PE_Check_Date = today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM PE1 Checked!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM PE1 Checked!" + System.Environment.NewLine + DSUM.Remark;
                        }
                        if (CurrentStatusID == 5) // W-JP_PAE-Check
                        {
                            DSUM.JP_PAE_PIC = name + "/" + today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM JP_PAE Checked!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM JP_PAE Checked!" + System.Environment.NewLine + DSUM.Remark;
                            die.DFM_PAE_Check_Date = today.ToString("yyyy/MM/dd");
                        }
                        if (CurrentStatusID == 6) // W-JP_PE1-Check
                        {
                            DSUM.JP_PE1_PIC = name + "/" + today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM JP_PE1 Checked!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM JP_PE1 Checked!" + System.Environment.NewLine + DSUM.Remark;
                            die.DFM_PE_Check_Date = today.ToString("yyyy/MM/dd");
                        }
                        if (CurrentStatusID == 7) // W-PE1-G6UP-Approve
                        {
                            DSUM.PE1M1AppBy = name + "/" + today.ToString("yyyy/MM/dd");
                            DSUM.PE1M1AppDate = today;
                            die.DFM_PE_App_Date = today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM PE1 Approved!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM PE1 Approved!" + System.Environment.NewLine + DSUM.Remark;

                        }
                        if (CurrentStatusID == 8) // W-PAE-G6UP-App
                        {
                            DSUM.PAEApproveBy = name + "/" + today.ToString("yyyy/MM/dd");
                            DSUM.PAEApproveDate = today;
                            die.DFM_PAE_App_Date = today.ToString("yyyy/MM/dd");
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM PAE Approved!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM PAE Approved!" + System.Environment.NewLine + DSUM.Remark;
                        }
                        if (CurrentStatusID == 12) // W-PE1-PIC-Check
                        {
                            DSUM.PE1CheckBy = name;
                            DSUM.PE1CheckDate = today;
                            die.Genaral_Information = today.ToString("yyyy/MM/dd") + ": DFM PE1 PIC Chhecked!" + System.Environment.NewLine + die.Genaral_Information;
                            DSUM.Remark = today.ToString("yyyy/MM/dd") + ": DFM PE1 PIC Chhecked!" + System.Environment.NewLine + DSUM.Remark;
                        }

                        if (CurrentStatusID == 13) // W-ECN-ISSe
                        {
                            DSUM.Remark = today.ToString("yyyy/MM/dd ") + name + ": Confirm ECN Issued!" + System.Environment.NewLine + DSUM.Remark;
                        }

                        DSUM.LastestReviseDate = today;

                        // Luu Attach DFM
                        //3.2 Luu Attachment
                        var fileExtCoverpage = Path.GetExtension(fileDFM.FileName);
                        if (!fileExtCoverpage.ToLower().Contains(".xls")) // file excel
                        {
                            return Json(new { status = false, msg = "File  must be excel" }, JsonRequestBehavior.AllowGet);
                        }
                        // Code luu file vào folder
                        string fileName = "";
                        if (DSUM.DSUMStatusID == 9) // Finished)
                        {
                            fileName = "[" + DSUM.DSUMNo + "]" + die.DieNo + fileExtCoverpage;
                            type = "DFM_Finished";
                        }
                        else
                        {
                            fileName = type + "_" + die.DieNo + today.ToString("yyyy-MM-dd-HHmmss") + fileExtCoverpage;
                        }

                        var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);



                        fileDFM.SaveAs(path);

                        // Đọc thông tin trong cover page
                        string controlLatestVersion = "By " + name + "_" + today.ToString("yyyy/MM/dd HH:mm:ss tt");
                        bool isSuccessSign = false;
                        if (today > applyDate)
                        {
                            die = readCoverpage(die, fileDFM);

                            isSuccessSign = getSignatureAndSaveCoverPage(path, CurrentStatusID, name, false, false, "", controlLatestVersion, false, null, die.Die_Code);

                            if (isSuccessSign)
                            {
                                DSUM.ControlLastestDFM = controlLatestVersion;
                            }
                            else
                            {
                                goto Next;
                            }
                        }
                        // covert to PDF

                        bool isSuccessCover = handelDSMtoPDF(fileName);
                        if (isSuccessCover == false)
                        {
                            status = false;
                            msg = "Coverpage is not insert(Embed) DFM Content";
                            goto Next;
                        }
                        Attachment newAtt = new Attachment()
                        {
                            DieID = die.DieID,
                            DFMID = DSUM.DFMID,
                            FileName = fileName,
                            Clasify = type,
                            CreateBy = Session["Name"].ToString(),
                            CreateDate = today,
                        };


                        db.Attachments.Add(newAtt);
                        db.SaveChanges();
                    }

                    if (!String.IsNullOrWhiteSpace(remark))
                    {
                        DSUM.Remark = today.ToString("MM/dd/yyyy ") + name + " : " + remark + System.Environment.NewLine + DSUM.Remark;
                    }

                    if (currentIndex == routeArray.Length - 1)
                    {
                        DSUM.DSUMStatusID = 9; // Finished

                    }
                    else
                    {
                        DSUM.DSUMStatusID = int.Parse(routeArray[currentIndex + 1]); // New status
                    }

                    var newStatus = db.DSUMStatusCategories.Find(DSUM.DSUMStatusID);

                    db.Entry(DSUM).State = EntityState.Modified;
                    db.Entry(die).State = EntityState.Modified;
                    db.SaveChanges();
                    status = true;
                    mailJob.sendEmailCheckDFM(DSUM, false, "");
                }
                else // No permit
                {
                    msg = "You do not have permition!";
                    status = false;
                }
            }
            else
            {

                msg = "Error!, Not enough data";
                status = false;
            }

        Next:
            var output = new
            {
                status = status,
                msg = msg
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }


        public JsonResult ReviseDFM(int attachID, string attachType, string reviseContent, HttpPostedFileBase fileDFM, string nextRouteID)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["DSUM_Role"].ToString();
            var statusConfig = db.DSUMStatusCategories.Where(x => x.StatusDone.Contains(attachType)).FirstOrDefault();
            bool verify = false;





            if (attachType == "DFM_SUPPLIER_SUBMIT")
            {
                verify = true;
            }
            else
            {
                if ((dept.Contains(statusConfig.DeptResponse) || statusConfig.DeptResponse.Contains(dept)) && (statusConfig.RoleResponse == role || role == "Approve"))
                {
                    verify = true;
                }
                else
                {
                    // No Permit
                    return Json(new { status = false, msg = "You no permit to revise this. Please re-check you role" }, JsonRequestBehavior.AllowGet);
                }
            }

            // code after verify permit OK
            var msg = "";
            var status = false;
            var name = Session["name"].ToString();
            var today = DateTime.Now;
            var att = db.Attachments.Find(attachID);
            var die = db.Die1.Find(att.DieID);
            var DSUM = db.DSUMs.Find(att.DFMID);

            /// Nâng version nếu bản revise đã final approved
            bool isUpver = false;
            int newVer = 0;
            if (DSUM.DSUMStatusID == 9)
            {
                DSUM.DSUMNo = genarateDSUMNo(DSUM.DSUMNo, false, true);
                isUpver = true;
                newVer = int.Parse(DSUM.DSUMNo.Substring(DSUM.DSUMNo.Length - 2, 2));
            }

            //check form  
            verify verifyResult = verifyFormCoverPage(fileDFM, DSUM.PartNo, DSUM.ControlLastestDFM, false);
            if (verifyResult.isPass == false)
            {
                msg = verifyResult.msg;
                status = false;
                goto Next;
            }
            //***

            DSUM.LastestReviseDate = today;
            DSUM.Remark = today.ToString("MM/dd/yyyy ") + name + " : " + "Revised DFM" + System.Environment.NewLine + reviseContent + System.Environment.NewLine + DSUM.Remark;
            // New status after revised
            string[] routeArray = DSUM.Route?.Split(',');
            int currentIndexRevise = -1;
            if (routeArray == null)
            {
                DSUM.DSUMStatusID = 1; // W-meeting
            }
            else
            {
                currentIndexRevise = Array.IndexOf(routeArray, statusConfig.DSUMStatusID.ToString());

                if (currentIndexRevise == routeArray.Length - 1)
                {
                    DSUM.DSUMStatusID = 9; // Finished
                }
                else
                {
                    DSUM.DSUMStatusID = int.Parse(nextRouteID); // New status
                }
            }

            //1. Luu Attachment

            var fileExt = Path.GetExtension(fileDFM.FileName);
            if (!fileExt.ToLower().Contains(".xls")) // file excel
            {
                return Json(new { status = false, msg = "File  must be excel" }, JsonRequestBehavior.AllowGet);
            }
            // Code luu file vào folder
            string fileName = "";
            if (DSUM.DSUMStatusID == 9) // Finished)
            {
                fileName = "[" + DSUM.DSUMNo + "]" + die.DieNo + fileExt;
            }
            else
            {
                fileName = "[Revised]" + att.Clasify + "_" + die.DieNo + today.ToString("yyyy-MM-dd-HHmmss") + fileExt;
            }
            var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
            fileDFM.SaveAs(path);

            att.FileName = fileName;
            att.ReviseBy = name;
            att.ReviseDate = DateTime.Now;
            // Đọc thông tin trong cover page

            string controlLatestVersion = "By " + name + "_" + today.ToString("yyyy/MM/dd HH:mm:ss tt");
            bool isSuccessSign = false;
            if (today > applyDate)
            {
                die = readCoverpage(die, fileDFM);
                isSuccessSign = getSignatureAndSaveCoverPage(path, statusConfig.DSUMStatusID, name, true, false, reviseContent, controlLatestVersion, isUpver, newVer, die.Die_Code);
                if (isSuccessSign)
                {
                    DSUM.ControlLastestDFM = controlLatestVersion;
                }

            }


            bool isSuccessCover = handelDSMtoPDF(fileName);
            if (isSuccessCover == false)
            {
                status = false;
                msg = "Coverpage is not insert(Embed) DFM Content";
                goto Next;
            }
            db.Entry(DSUM).State = EntityState.Modified;


            db.Entry(att).State = EntityState.Modified;
            db.Entry(die).State = EntityState.Modified;
            db.SaveChanges();
            mailJob.sendEmailCheckDFM(DSUM, true, reviseContent);
            status = true;

        Next:
            var output = new
            {
                status = status,
                msg = msg,
                DSUMID = DSUM.DFMID
            };

            // mailJob.sendEmailCheckDFM(reciever, sender, mailmsg, item, deptRecive, die.SupplierID != 0 ? die.Supplier.SupplierName : "N/A", die.ModelList.ModelName, false);

            return Json(output, JsonRequestBehavior.AllowGet);


        }


        public JsonResult getCurrentRoute(int attachID, string attachType)
        {
            var att = db.Attachments.Find(attachID);
            var DSUM = db.DSUMs.Find(att.DFMID);
            var status = false;
            var msg = "";
            List<object> remainRoute = new List<object>();
            if (attachType == "DFM_SUPPLIER_SUBMIT")
            {
                remainRoute.Add(new
                {
                    route = new
                    {
                        RouteID = 1,
                        RouteName = "W-Meeting"
                    }
                });
                status = true;
            }
            else
            {
                var statusConfig = db.DSUMStatusCategories.Where(x => x.StatusDone.Contains(attachType)).FirstOrDefault();
                string[] routeArray = DSUM.Route?.Split(',');
                var currentStatus = DSUM.DSUMStatusID.ToString();
                int currentIndexRevise = Array.IndexOf(routeArray, statusConfig.DSUMStatusID.ToString());
                int currentStatusIndex = Array.IndexOf(routeArray, currentStatus);
                if (currentIndexRevise == -1 && attachType != "DFM After Meeting")
                {
                    status = false;
                    msg = "You revise DFM but it not include current Route";
                }
                else
                {
                    currentIndexRevise = currentIndexRevise == -1 ? 0 : currentIndexRevise;
                    for (int i = currentIndexRevise + 1; i < currentStatusIndex + 1; i++)
                    {
                        remainRoute.Add(new
                        {
                            route = new
                            {
                                RouteID = int.Parse(routeArray[i]),
                                RouteName = db.DSUMStatusCategories.Find(int.Parse(routeArray[i])).RouteName
                            }
                        });

                    }
                    status = true;
                }

            }




            return Json(new { status = status, remainRoute = remainRoute, msg = msg }, JsonRequestBehavior.AllowGet);
        }


        public JsonResult cancelDFM(int DFMID, string reason)
        {

            var dept = Session["Dept"].ToString();
            var role = Session["Die_Lauch_Role"].ToString();
            if (role != "Edit" || !dept.Contains("PAE"))
            {
                return Json(new { status = false, msg = "No permision!" }, JsonRequestBehavior.AllowGet);
            }

            var status = false;
            var DSUM = db.DSUMs.Find(DFMID);
            DSUM.DSUMStatusID = 11; // cancel
            var die = db.Die1.Where(x => x.PartNoOriginal == DSUM.PartNo && x.Die_Code == DSUM.DieNo && x.Active != false && x.isCancel != true).FirstOrDefault();
            die.DFM_Sub_Date = null;
            die.DFM_PAE_Check_Date = null;
            die.DFM_PE_Check_Date = null;
            die.DFM_PE_App_Date = null;
            die.DFM_PAE_App_Date = null;
            DSUM.LastestReviseDate = DateTime.Now;
            DSUM.Remark = DateTime.Now.ToString("MM/dd/yyyy ") + Session["Name"] + " : " + "Cancel DFM: " + reason + System.Environment.NewLine + DSUM.Remark;
            db.Entry(DSUM).State = EntityState.Modified;
            db.Entry(die).State = EntityState.Modified;
            db.SaveChanges();
            status = true;
            var output = new
            {
                status = status,
                msg = "OK"
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getSumarize()
        {
            var WPAEPICCheck = db.DSUMs.Where(x => x.DSUMStatusID == 1).Count();
            var WPE1PICCheck = db.DSUMs.Where(x => x.DSUMStatusID == 2).Count();
            var WPE1Approve = db.DSUMs.Where(x => x.DSUMStatusID == 3).Count();
            var WPAEApprove = db.DSUMs.Where(x => x.DSUMStatusID == 4).Count();
            var Rejected = db.DSUMs.Where(x => x.DSUMStatusID == 6).Count();
            var Cancelled = db.DSUMs.Where(x => x.DSUMStatusID == 7).Count();
            var Finhished = db.DSUMs.Where(x => x.DSUMStatusID == 5).Count();

            var output = new
            {
                WPAEPICCheck = WPAEPICCheck,
                WPE1PICCheck = WPE1PICCheck,
                WPE1Approve = WPE1Approve,
                WPAEApprove = WPAEApprove,
                Rejected = Rejected,
                Cancelled = Cancelled,
                Finhished = Finhished
            };
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public Die1 readCoverpage(Die1 die, HttpPostedFileBase coverpage)
        {
            try
            {
                if (coverpage == null)
                {
                    return die;
                }
                var today = DateTime.Now;
                using (ExcelPackage package = new ExcelPackage(coverpage.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    //*************************************************************
                    //*************************************************************
                    //*************************************************************
                    //*************************************************************
                    // Read phần common
                    //0. MO OR PX
                    string field = worksheet.Cells["A1"].Text?.Trim().ToUpper();
                    bool isMofield = field.Contains("MO");
                    //1. Family
                    var familyPart = worksheet.Cells["F17"].Text?.Trim().ToUpper();
                    if (!String.IsNullOrWhiteSpace(familyPart))
                    {
                        string out_Fami = "";

                        string[] listFamily = familyPart.Split(',');
                        foreach (var item in listFamily)
                        {
                            if (item.Trim().ToUpper().Length == 8)
                            {
                                out_Fami = out_Fami + "," + item.Trim().ToUpper() + "-000";
                            }
                            else
                            {
                                out_Fami = out_Fami + "," + item.Trim().ToUpper();
                            }
                        }
                        die.Family_Die_With = out_Fami;
                    }

                    //2. Common
                    var commonPart = worksheet.Cells["F18"].Text?.Trim().ToUpper();
                    if (!String.IsNullOrWhiteSpace(commonPart))
                    {
                        string out_Comm = "";

                        string[] listComm = commonPart.Split(',');
                        foreach (var item in listComm)
                        {
                            if (item.Trim().ToUpper().Length == 8)
                            {
                                out_Comm = out_Comm + "," + item.Trim().ToUpper() + "-000";
                            }
                            else
                            {
                                out_Comm = out_Comm + "," + item.Trim().ToUpper();
                            }
                        }
                        die.Common_Part_With = out_Comm;
                    }

                    //3. Cav qty
                    {
                        int? cavQty = 0;
                        string cavJud = worksheet.Cells["H21"].Text?.Trim().ToUpper();
                        if (cavJud.Contains("REVISE"))
                        {
                            cavQty = commonFunction.getNummberInString(worksheet.Cells["I21"].Text?.Trim().ToUpper());
                            //int.TryParse(worksheet.Cells["I21"].Text?.Trim().ToUpper(), out cavQty);
                        }
                        else
                        {
                            cavQty = commonFunction.getNummberInString(worksheet.Cells["F21"].Text?.Trim().ToUpper());
                            //int.TryParse(worksheet.Cells["F21"].Text?.Trim().ToUpper(), out cavQty);
                        }
                        die.CavQuantity = cavQty;
                    }

                    //4. MC size
                    {
                        int? mcSize = 0;
                        string MCJud = worksheet.Cells["H22"].Text?.Trim().ToUpper();
                        if (MCJud.Contains("REVISE"))
                        {
                            mcSize = commonFunction.getNummberInString(worksheet.Cells["I22"].Text?.Trim().ToUpper());
                            //int.TryParse(worksheet.Cells["I22"].Text?.Trim().ToUpper(), out mcSize);
                        }
                        else
                        {
                            mcSize = commonFunction.getNummberInString(worksheet.Cells["F22"].Text?.Trim().ToUpper());
                            // int.TryParse(worksheet.Cells["F22"].Text?.Trim().ToUpper(), out mcSize);
                        }
                        die.MCsize = mcSize;
                    }

                    // 5. Idea sumarize
                    string sumIdea = "";
                    int Q = 0;
                    int C = 0;
                    int D = 0;
                    for (int row = 33; row <= 40; row++)
                    { // Row by row...
                        string idea = worksheet.Cells["R" + row.ToString()].Text?.Trim();
                        string forWhat = worksheet.Cells["AD" + row.ToString()].Text?.Trim().ToUpper();
                        if (!String.IsNullOrWhiteSpace(idea))
                        {
                            sumIdea += System.Environment.NewLine + idea + " for " + forWhat;
                            Q = Q + (forWhat.Contains("Q") ? 1 : 0);
                            C = C + (forWhat.Contains("C") ? 1 : 0);
                            D = D + (forWhat.Contains("D") ? 1 : 0);
                        }
                    }
                    die.DSUM_Idea = sumIdea;
                    die.NOofForQ = Q;
                    die.NOofForC = C;
                    die.NOofForD = D;

                    //************************************************************
                    //************************************************************
                    // Read phần MO
                    if (isMofield)
                    {
                        //5. Gate
                        {
                            string gate = "";
                            string jude = worksheet.Cells["H23"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                gate = worksheet.Cells["I23"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                gate = worksheet.Cells["F23"].Text?.Trim().ToUpper();
                            }
                            die.GateType = gate;
                        }

                        //6. HR
                        {
                            string hr = "";
                            string jude = worksheet.Cells["H24"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                hr = worksheet.Cells["I24"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                hr = worksheet.Cells["F24"].Text?.Trim().ToUpper();
                            }
                            die.HotRunner = hr;

                        }

                        //7. Core/Cav material
                        {
                            string value = "";
                            string jude = worksheet.Cells["H29"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I29"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F29"].Text?.Trim().ToUpper();
                            }
                            die.CoreCavMaterial = value;
                        }

                        //8. slider material
                        {
                            string value = "";
                            string jude = worksheet.Cells["H30"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I30"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F30"].Text?.Trim().ToUpper();
                            }
                            die.SliderMaterial = value;
                        }
                        //9. Lifter material
                        {
                            string value = "";
                            string jude = worksheet.Cells["H31"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I31"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F31"].Text?.Trim().ToUpper();
                            }
                            die.LifterMaterial = value;
                        }

                        // Warranty
                        {
                            double value = 0;
                            string jude = worksheet.Cells["H34"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {

                                double.TryParse(worksheet.Cells["I34"].Text?.Trim().ToUpper(), out value);
                            }
                            else
                            {
                                double.TryParse(worksheet.Cells["F34"].Text?.Trim().ToUpper(), out value);
                            }
                            die.WarrantyShotAsDSUM = value * 1000000;
                        }

                        //10. Die Maker
                        {
                            string value = "";
                            string jude = worksheet.Cells["H37"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I37"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F37"].Text?.Trim().ToUpper();
                            }
                            die.DieMaker = value;
                        }

                        //11. Die Make locatoion
                        {
                            string value = "";
                            string jude = worksheet.Cells["H38"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I38"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F38"].Text?.Trim().ToUpper();
                            }
                            die.DieMakeLocation = value;
                        }

                        //11. Part Material
                        {
                            string value1 = "";
                            string jude1 = worksheet.Cells["O21"].Text?.Trim().ToUpper();
                            if (jude1.Contains("REVISE"))
                            {
                                value1 = worksheet.Cells["P21"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value1 = worksheet.Cells["M21"].Text?.Trim().ToUpper();
                            }

                            //***
                            string value2 = "";
                            string jude2 = worksheet.Cells["O22"].Text?.Trim().ToUpper();
                            if (jude2.Contains("REVISE"))
                            {
                                value2 = worksheet.Cells["P22"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value2 = worksheet.Cells["M22"].Text?.Trim().ToUpper();
                            }


                            var part = db.Parts1.Where(x => x.PartNo == die.PartNoOriginal && x.Active != false).FirstOrDefault();
                            if (part != null)
                            {
                                part.Material = value1 + "/" + value2;
                                db.Entry(part).State = EntityState.Modified;
                                db.SaveChanges();
                            }


                        }

                        //12. Lead time PO-> TO
                        {
                            int? value = 0;
                            string jude = worksheet.Cells["H40"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = commonFunction.getNummberInString(worksheet.Cells["I40"].Text?.Trim().ToUpper());
                                //int.TryParse(worksheet.Cells["I40"].Text?.Trim().ToUpper(), out value);
                            }
                            else
                            {
                                value = commonFunction.getNummberInString(worksheet.Cells["F40"].Text?.Trim().ToUpper());
                                // int.TryParse(worksheet.Cells["F40"].Text?.Trim().ToUpper(), out value);
                            }
                            die.LeadTimeMakeDie = value;
                        }

                        //13. Texture
                        {
                            string value = "";
                            string jude = worksheet.Cells["O24"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["P24"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["M24"].Text?.Trim().ToUpper();
                            }

                            string[] configNoneTexture = { "NA", "N/A", "N\\A", "-", "--", "---", "", "O", "NO", "N" };
                            if (configNoneTexture.Contains(value))
                            {
                                die.Texture = false;
                            }
                            else
                            {
                                die.Texture = true;
                                die.TextureType = value;

                            }
                        }

                        ////14. JIG
                        //{

                        //    string value = worksheet.Cells["AG29"].Text?.Trim().ToUpper();
                        //    if (value == "1")
                        //    {
                        //        die.JIG_Using = true;
                        //    }
                        //    else
                        //    {
                        //        if (value == "2")
                        //        {
                        //            die.JIG_Using = false;
                        //        }
                        //    }
                        //}

                        ////15. New master gear
                        //{

                        //    string value = worksheet.Cells["AG32"].Text?.Trim().ToUpper();
                        //    if (value == "1")
                        //    {
                        //        die.isNewMasterGear = true;
                        //    }
                        //    else
                        //    {
                        //        if (value == "2")
                        //        {
                        //            die.isNewMasterGear = false;
                        //        }
                        //    }
                        //}

                        //16. High CT
                        {

                            string value = worksheet.Cells["AG35"].Text?.Trim().ToUpper();
                            string spec = "High Cycle Die, ";
                            if (value == "1")
                            {

                                if (die.SpecialSpec != null)
                                {
                                    if (!die.SpecialSpec.Contains(spec))
                                    {
                                        die.SpecialSpec = spec + die.SpecialSpec;
                                    }
                                }
                                else
                                {
                                    die.SpecialSpec = spec;
                                }
                            }
                            else
                            {
                                if (die.SpecialSpec != null)
                                {
                                    if (!die.SpecialSpec.Contains(spec))
                                    {
                                        die.SpecialSpec = die.SpecialSpec.Replace(spec, " ");
                                    }
                                }
                                else
                                {
                                    die.SpecialSpec = spec;
                                }
                            }

                        }

                        //17. Cycle Time
                        {
                            int? value = 0;
                            value = commonFunction.getNummberInString(worksheet.Cells["M31"].Text?.Trim().ToUpper());
                            //int.TryParse(worksheet.Cells["M31"].Text?.Trim().ToUpper(), out value);

                            die.CycleTime_Sec = value;
                        }
                        //17. Cycle Time target
                        {
                            int? value = 0;
                            value = commonFunction.getNummberInString(worksheet.Cells["M32"].Text?.Trim().ToUpper());
                            //int.TryParse(worksheet.Cells["M32"].Text?.Trim().ToUpper(), out value);
                            die.CycleTime_TargetAsDSUM = value;
                        }



                    }
                    else //   // Read phần PX
                    {
                        // Warranty
                        {
                            double value = 0;
                            string jude = worksheet.Cells["H23"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                double.TryParse(worksheet.Cells["I23"].Text?.Trim().ToUpper(), out value);
                            }
                            else
                            {
                                double.TryParse(worksheet.Cells["F23"].Text?.Trim().ToUpper(), out value);
                            }
                            die.WarrantyShotAsDSUM = value * 1000000;
                        }
                        // Die Maker
                        {
                            string value = "";
                            string jude = worksheet.Cells["H24"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I24"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F24"].Text?.Trim().ToUpper();
                            }
                            die.DieMaker = value;
                        }

                        // Die make location
                        {
                            string value = "";
                            string jude = worksheet.Cells["H25"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I25"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F25"].Text?.Trim().ToUpper();
                            }
                            die.DieMakeLocation = value;
                        }
                        //12. Lead time PO-> TO
                        {
                            int? value = 0;
                            string jude = worksheet.Cells["H29"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = commonFunction.getNummberInString(worksheet.Cells["I29"].Text?.Trim().ToUpper());
                                //int.TryParse(worksheet.Cells["I29"].Text?.Trim().ToUpper(), out value);
                            }
                            else
                            {
                                value = commonFunction.getNummberInString(worksheet.Cells["F29"].Text?.Trim().ToUpper());
                                //int.TryParse(worksheet.Cells["F29"].Text?.Trim().ToUpper(), out value);
                            }
                            die.LeadTimeMakeDie = value;
                        }

                        // Die Componant
                        {
                            string value = "";
                            string jude = worksheet.Cells["H32"].Text?.Trim().ToUpper();
                            if (jude.Contains("REVISE"))
                            {
                                value = worksheet.Cells["I32"].Text?.Trim().ToUpper();
                            }
                            else
                            {
                                value = worksheet.Cells["F32"].Text?.Trim().ToUpper();
                            }
                            die.PXNoOfComponent = value;
                        }

                        //// JIG
                        {

                            string value = worksheet.Cells["AG29"].Text?.Trim().ToUpper();
                            if (value == "1")
                            {
                                die.JIG_Using = true;
                            }
                            else
                            {
                                if (value == "2")
                                {
                                    die.JIG_Using = false;
                                }
                            }
                        }

                        // Stacking
                        {

                            string value = worksheet.Cells["AK29"].Text?.Trim().ToUpper();
                            if (value == "1")
                            {
                                die.isStacking = true;
                            }
                            else
                            {
                                if (value == "2")
                                {
                                    die.isStacking = false;
                                }
                            }
                        }

                        // Shaft Kashime
                        {

                            string value = worksheet.Cells["AG32"].Text?.Trim().ToUpper();
                            if (value == "1")
                            {
                                die.isShaftKashime = true;
                            }
                            else
                            {
                                if (value == "2")
                                {
                                    die.isShaftKashime = false;
                                }
                            }
                        }

                        // Burring Kashime
                        {

                            string value = worksheet.Cells["AK32"].Text?.Trim().ToUpper();
                            if (value == "1")
                            {
                                die.isBurringKashime = true;
                            }
                            else
                            {
                                if (value == "2")
                                {
                                    die.isBurringKashime = false;
                                }
                            }
                        }

                        // Progresive SPM
                        {

                            int value = 0;
                            int.TryParse(worksheet.Cells["P26"].Text?.Trim().ToUpper(), out value);

                            die.SPM_Progressive = value;
                        }

                        // singer SPM
                        {

                            int value = 0;
                            int.TryParse(worksheet.Cells["P27"].Text?.Trim().ToUpper(), out value);

                            die.SPM_Single = value;
                        }

                        //Die material
                        {
                            die.PuchHolder = worksheet.Cells["R54"].Text?.Trim().ToUpper();
                            die.PunchBackingPlate = worksheet.Cells["R57"].Text?.Trim().ToUpper();
                            die.PunchPlate = worksheet.Cells["R60"].Text?.Trim().ToUpper();
                            die.StripperBackingPlate = worksheet.Cells["R63"].Text?.Trim().ToUpper();
                            die.StripperPlate = worksheet.Cells["R66"].Text?.Trim().ToUpper();
                            die.DiePlate = worksheet.Cells["R69"].Text?.Trim().ToUpper();
                            die.DieBackingPlate = worksheet.Cells["R72"].Text?.Trim().ToUpper();
                            die.DieHolder = worksheet.Cells["R75"].Text?.Trim().ToUpper();
                            die.GetaPlate = worksheet.Cells["R78"].Text?.Trim().ToUpper();
                            die.Punch = worksheet.Cells["R84"].Text?.Trim().ToUpper();
                            die.InsertBlock = worksheet.Cells["R87"].Text?.Trim().ToUpper();

                        }

                    }


                    //************************************************************
                    //************************************************************

                }
            }
            catch
            {

            }


            return die;
        }

        public bool getSignatureAndSaveCoverPage(string fullCoverPath, int CurrentStatusID, string userName, bool isRevise, bool isNewDFM, string reviseContent, string controlLatest, bool isUpVer, int? newVer, string dieCode)
        {
            bool status = false;
            var today = DateTime.Now;
            try
            {
                using (ExcelPackage package = new ExcelPackage(fullCoverPath))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    // Update correct DIE NO on cover page
                    worksheet.Cells["M16"].Value = dieCode;

                    if (isRevise) // Truong hop revise
                    {
                        worksheet.Cells["Q42"].Value = "● " + today.ToString("yyyy/MM/dd ") + userName + " Revised: " + System.Environment.NewLine + reviseContent + System.Environment.NewLine + worksheet.Cells["Q42"].Text;
                        if (isUpVer == true)
                        {
                            for (int i = 2; i < 8; i++)
                            {
                                string[] configCellInputVer = { "-", "-", "AH26", "AI26", "AJ26", "AK26", "AL26", "AM26", "AN26" };
                                string[] configCellInputDate = { "-", "-", "I5", "K5", "M5", "O5", "Q5", "S5", "U5" };
                                if (i == newVer + 1)
                                {
                                    worksheet.Cells[configCellInputVer[i]].Value = true;
                                    worksheet.Cells[configCellInputDate[i]].Value = DateTime.Now.ToString("MM/dd/yyyy");
                                }
                            }

                        }
                    }
                    else
                    {
                        if (!isNewDFM)
                        {
                            if (CurrentStatusID == 1) // W-meeting 
                            {
                                // PAE PIC
                                worksheet.Cells["I9"].Value = userName + "\r\n" + today.ToString("dd-MMM-yy");

                            }
                            if (CurrentStatusID == 2) // W-PAE-G4UP-Check
                            {
                                // PAE G4UP
                                string currentI13 = worksheet.Cells["I13"].Text;

                                worksheet.Cells["I13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy") + (String.IsNullOrWhiteSpace(currentI13) || currentI13 == "\r\n" ? "" : Environment.NewLine + currentI13);
                            }
                            if (CurrentStatusID == 3) // W-DMT-G6UP-Check
                            {
                                // DMT G6UP
                                worksheet.Cells["E13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy");

                            }
                            if (CurrentStatusID == 4) // W-PE1- G4UP-Check
                            {
                                // PE1 G4UP

                                string currentQ13 = worksheet.Cells["Q13"].Text;
                                worksheet.Cells["Q13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy") + (String.IsNullOrWhiteSpace(currentQ13) || currentQ13 == "\r\n" ? "" : Environment.NewLine + currentQ13);
                            }
                            if (CurrentStatusID == 5) // W-JP_PAE-Check
                            {
                                // PAE G4UP
                                string currentI13 = worksheet.Cells["I13"].Text;
                                worksheet.Cells["I13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy") + (String.IsNullOrWhiteSpace(currentI13) || currentI13 == "\r\n" ? "" : Environment.NewLine + currentI13);

                            }
                            if (CurrentStatusID == 6) // W-JP_PE1-Check
                            {
                                // PE1 G4UP

                                string currentQ13 = worksheet.Cells["Q13"].Text;
                                worksheet.Cells["Q13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy") + (String.IsNullOrWhiteSpace(currentQ13) || currentQ13 == "\r\n" ? "" : Environment.NewLine + currentQ13);
                            }
                            if (CurrentStatusID == 7) // W-PE1-G6UP-Approve
                            {

                                worksheet.Cells["T13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy");
                            }
                            if (CurrentStatusID == 8) // W-PAE-G6UP-App
                            {
                                worksheet.Cells["M13"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy");
                            }
                            if (CurrentStatusID == 12) // W-PE1-PIC-Check
                            {
                                worksheet.Cells["M9"].Value = userName + Environment.NewLine + today.ToString("dd-MMM-yy");
                            }
                        }
                        else
                        {
                            worksheet.Cells["AG26"].Value = true;
                            worksheet.Cells["G5"].Value = DateTime.Now.ToString("MM/dd/yyyy");

                        }


                    }

                    // add controlLastes
                    worksheet.Cells["A4"].Value = controlLatest;
                    package.Workbook.Calculate();
                    package.Save();
                    status = true;
                }
            }
            catch
            {
                status = false;
            }

            return status;
        }
        public void testInsertObj()
        {
            var path = Path.Combine(Server.MapPath("~/File/Attachment/"), "test2.xlsx");
            var path2 = Path.Combine(Server.MapPath("~/File/Attachment/"), "BOX.pptx_RC5-2020-000-21A-0012023-06-21-112818.pptx");
            string srcImage = Path.Combine(Server.MapPath("~/File/Attachment/"), "DFMContentDefaulImage.png");

            Bitmap bitmap = new Bitmap(srcImage);
            byte[] imageData = null;
            byte[] fileEmbed = null;

            using (FileStream fs = new FileStream(srcImage, FileMode.Open, FileAccess.Read))
            {
                // Create a byte array of file stream length
                imageData = System.IO.File.ReadAllBytes(srcImage);
                //Read block of bytes from stream into the byte array
                fs.Read(imageData, 0, System.Convert.ToInt32(fs.Length));
                //Close the File Stream
                fs.Close();
            }

            using (FileStream fs = new FileStream(path2, FileMode.Open, FileAccess.Read))
            {
                // Create a byte array of file stream length
                fileEmbed = System.IO.File.ReadAllBytes(path2);
                //Read block of bytes from stream into the byte array
                fs.Read(fileEmbed, 0, System.Convert.ToInt32(fs.Length));
                //Close the File Stream
                fs.Close();
            }


            using (Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(path))
            {
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Shapes.AddOleObject(4, 0, 5, 0, 100, 200, imageData);
                // Set embedded ole object data.
                worksheet.OleObjects[0].ObjectData = fileEmbed;
                worksheet.OleObjects[0].DisplayAsIcon = true;
                workbook.Save(path, Aspose.Cells.SaveFormat.Xlsx);

            };



            //using (Spire.Xls.Workbook workbook = new Spire.Xls.Workbook())
            //{
            //    workbook.LoadFromFile(path, ExcelVersion.Version2007);
            //    Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
            //    Spire.Xls.Core.IOleObject oleObject = sheet.OleObjects.Add(path2, bitmap, OleLinkType.Embed);
            //    oleObject.Location = sheet.Range["B4:C7"];
            //    oleObject.DisplayAsIcon = true;
            //    workbook.Save();


            //}
        }

        public bool handelDSMtoPDF(string fileCoverName)
        {
            // fileCoverName = "[Rev4-11-180807.xlsx";

            bool status = false;
            try
            {
                string rundomString = Guid.NewGuid().ToString();
                var fullPathCoverPage = Path.Combine(Server.MapPath("~/File/Attachment/"), fileCoverName);
                var pathTemp = Server.MapPath("~/File/Attachment/Temp/");
                string fileName = "FileDFMContent_" + rundomString + "_" + ".";
                var fullPathCoverpagePDF = "";
                var fullPathContentPDF = "";

                //****** Get file OBJ được insert trong coverpage => luu thành file PDF
                using (Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(fullPathCoverPage))
                {
                    Aspose.Cells.Worksheet worksheet = workbook.Worksheets.First();
                    // Get the OleObject Collection in the first worksheet.
                    Aspose.Cells.Drawing.OleObjectCollection oles = workbook.Worksheets[0].OleObjects;

                    // Loop through all the oleobjects and extract each object in the worksheet.
                    if (oles.Count() == 0)
                    {
                        return false;
                    }

                    Aspose.Cells.Drawing.OleObject ole = oles[0];
                    // Specify the output filename.

                    string mainfileName = fileName;
                    // Specify each file format based on the oleobject format type.
                    string formatType = ole.FileFormatType.ToString().ToLower();
                    fileName += formatType;


                    // Set active sheet to output
                    var cateID = worksheet.Cells["AG18"].Value != null ? worksheet.Cells["AG18"].Value.ToString() : "";

                    bool isMofield = worksheet.Cells["A1"].Value.ToString().Contains("MO");




                    if (cateID != "1")
                    {
                        ////*** Tạm thời cover cũ
                        //if (cateID == "")
                        //{
                        //    if (isMofield)
                        //    {
                        //        worksheet.PageSetup.PrintArea = "A1:Z46";

                        //    }
                        //    else
                        //    {
                        //        worksheet.PageSetup.PrintArea = "A1:AA93";
                        //    }
                        //}
                        //else
                        {
                            if (isMofield)
                            {
                                worksheet.PageSetup.PrintArea = "A1:AE91";

                            }
                            else
                            {
                                worksheet.PageSetup.PrintArea = "A1:AE136";
                            }
                        }


                    }
                    else
                    {


                        if (isMofield)
                        {
                            worksheet.PageSetup.PrintArea = "A1:AE46";

                        }
                        else
                        {
                            worksheet.PageSetup.PrintArea = "A1:AE91";
                        }


                    }
                    // saveOption Aspose.Cells

                    worksheet.PageSetup.Orientation = Aspose.Cells.PageOrientationType.Landscape;
                    worksheet.PageSetup.PaperSize = Aspose.Cells.PaperSizeType.PaperA4;
                    workbook.Save(Path.Combine(pathTemp + rundomString + fileCoverName.Split('.')[0] + ".pdf"), Aspose.Cells.SaveFormat.Pdf);
                    fullPathCoverpagePDF = Path.Combine(pathTemp + rundomString + fileCoverName.Split('.')[0] + ".pdf");
                    fullPathCoverpagePDF = editPDFtoRemoveLisenceText(pathTemp + rundomString + fileCoverName.Split('.')[0] + ".pdf");
                    // Save the oleobject as a new excel file if the object type is xls. => PDF
                    if (ole.FileFormatType == FileFormatType.Xlsx)
                    {
                        MemoryStream ms = new MemoryStream();
                        if (ole.ObjectData != null)
                        {
                            ms.Write(ole.ObjectData, 0, ole.ObjectData.Length);
                            Aspose.Cells.Workbook oleBook = new Aspose.Cells.Workbook(ms);
                            oleBook.Settings.IsHidden = false;
                            oleBook.Save(Path.Combine(pathTemp + fileName), Aspose.Cells.SaveFormat.Pdf);
                            // fullPathContentPDF = pathTemp + mainfileName + ".pdf";
                        }
                    }
                    // Create the files based on the oleobject format types.
                    else
                    {
                        if (ole.ObjectData != null)
                        {
                            FileStream fs = System.IO.File.Create(pathTemp + fileName);
                            fs.Write(ole.ObjectData, 0, ole.ObjectData.Length);
                            fs.Close();
                            fullPathContentPDF = converttoPDF(pathTemp + fileName);
                        }
                    }
                }

                // Nối 2 file PDF
                string newPDFFileName = "[Merge]" + fileCoverName.Split('.')[0] + ".pdf";
                string fullOutPutPathMerge = Path.Combine(Server.MapPath("~/File/Attachment/Temp/"), newPDFFileName);
                string[] fileNeedMerge = { fullPathCoverpagePDF, fullPathContentPDF };
                string mergedFile = CombineMultiplePDFs(fileNeedMerge, fullOutPutPathMerge);
                string DFM_PDF_DoneFile = AddPageNoForPDF(mergedFile, Server.MapPath("~/File/Attachment/"));


                deleteFileStoreInTempFolder(fullPathCoverpagePDF);
                deleteFileStoreInTempFolder(fullPathContentPDF);
                deleteFileStoreInTempFolder(pathTemp + fileName);
                deleteFileStoreInTempFolder(mergedFile);
                status = true;
            }
            catch
            {
                status = false;
            }


            return status;

        }

        public class verify
        {
            public bool isPass { set; get; }
            public string msg { set; get; }
        }
        public verify verifyFormCoverPage(HttpPostedFileBase coverpage, string partNo, string latestVersionControl, bool isNew)
        {

            bool isPass = false;
            string msg = "";
            var today = DateTime.Now;
            if (today < applyDate)
            {
                isPass = true;
                msg = "";
                goto Exit;

            }

            if (coverpage == null)
            {
                isPass = false;
                msg = "No file Upload";
                goto Exit;
            }

            var fileExtCoverpage = Path.GetExtension(coverpage.FileName);
            if (!fileExtCoverpage.ToLower().Contains(".xls")) // file excel
            {
                isPass = false;
                msg = "File must be excel (.xls*)";
                goto Exit;
            }

            string MOversion = "LPE-0012- Att 01-Rev.08\nEffective date: Dec/01/2023";
            string PXversion = "LPE-0012- Att 01-Rev.09\nEffective date: Dec/01/2023";
            string A45 = "Auto by Sys";

            using (ExcelPackage package = new ExcelPackage(coverpage.InputStream))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                //1. Kiem tra form dung ko?
                string field = worksheet.Cells["A1"].Text?.Trim().ToUpper();

                bool isMofield = field.Contains("MO");
                string formVersion = worksheet.Cells["AB1"].Text.Trim();
                string a45 = worksheet.Cells["A45"].Text.Trim();
                string formPartNo = worksheet.Cells["F16"].Text.Trim().ToUpper();
                string latestDFMVersion = worksheet.Cells["A4"].Text.Trim();
                // Check version
                if (isMofield)
                {
                    if (formVersion != MOversion)
                    {
                        isPass = false;
                        msg = "Not correct version! form you upload is " + formVersion + " but sys request version " + MOversion;
                        goto Exit;
                    }
                }
                else
                {
                    if (formVersion != PXversion)
                    {
                        isPass = false;
                        msg = "Not correct version! form you upload is " + formVersion + " but sys request version " + PXversion;
                        goto Exit;
                    }
                }

                // Check dong A45 
                if (A45 != a45)
                {
                    isPass = false;
                    msg = "Not correct format! Maybe you add/delete row, Please compare with version " + (isMofield ? MOversion : PXversion);
                    goto Exit;
                }

                //2. Kiem tra part No trung ko?
                if (!partNo.Contains(formPartNo))
                {
                    isPass = false;
                    msg = "Maybe you upload wrong Part No. In coverpage PartNo is " + formPartNo + " Part No on DMS is " + partNo;
                    goto Exit;
                }


                //3. Kiem tra lastest version ko?
                if (latestDFMVersion != latestVersionControl && isNew == false)
                {
                    isPass = false;
                    msg = "DFM you upload is not newest version, newest was uploaded " + latestVersionControl;
                    goto Exit;
                }

                isPass = true;
                msg = "OK";

            }



        Exit:
            var output = new verify()
            {
                isPass = isPass,
                msg = msg
            };
            return output;

        }


        public string getRole(int DFMID)
        {
            string dept = Session["Dept"].ToString();
            string role = Session["DSUM_Role"].ToString();
            string grade = Session["Grade"].ToString();

            var DSUM = db.DSUMs.Find(DFMID);

            // Check role
            var statusConfig = db.DSUMStatusCategories.Find(DSUM.DSUMStatusID);

            string view = "JUST VIEW";
            if ((dept.Contains(statusConfig.DeptResponse) || statusConfig.DeptResponse == "All") && statusConfig.GradeResponse.Contains(grade) && (statusConfig.RoleResponse == role || role == "Approve"))
            {
                view = statusConfig.RoleResponse;
                if (DSUM.DSUMStatusID == 13)
                {
                    view = "ECN ISSUED";
                }
            }


            return view;
        }


        public string converttoPDF(string fullPath)
        {


            string outputFullPath = "";
            string fileExtention = Path.GetExtension(fullPath);
            string fileNameNoExtension = Path.GetFileNameWithoutExtension(fullPath);
            string currentpath = Path.GetDirectoryName(fullPath) + "\\";
            if (fileExtention.ToLower().Contains("ppt"))
            {
                PDFNet.Initialize("demo:1695198894695:7c01841303000000006be885da1dd188c75c9ed579b6b674c75bf5fe6c");

                // Start with a PDFDoc (the conversion destination)
                using (PDFDoc pdfdoc = new PDFDoc())
                {
                    // perform the conversion with no optional parameters
                    pdftron.PDF.Convert.OfficeToPDF(pdfdoc, fullPath, null);

                    // save the result
                    pdfdoc.Save(currentpath + "[PDF]" + fileNameNoExtension + ".pdf", SDFDoc.SaveOptions.e_linearized);
                    outputFullPath = currentpath + "[PDF]" + fileNameNoExtension + ".pdf";

                }
            }
            else
            {
                if (fileExtention.ToLower().Contains("pdf"))
                {
                    outputFullPath = fullPath;
                }
            }



            return outputFullPath;
        }


        public string testconverttoPDF(string path, string fileName)
        {
            string testPath_in = Server.MapPath("~/File/Attachment/Temp/RC5-4182-14A-DFM 13-Sep.pptx");
            string testPath_out = Server.MapPath("~/File/Attachment/Temp/[pdf]RC5-4182-14A-DFM 13-Sep.pdf");
            string msg = "";

            PDFNet.Initialize("demo:1695198894695:7c01841303000000006be885da1dd188c75c9ed579b6b674c75bf5fe6c");

            // Start with a PDFDoc (the conversion destination)
            using (PDFDoc pdfdoc = new PDFDoc())
            {
                // perform the conversion with no optional parameters
                pdftron.PDF.Convert.OfficeToPDF(pdfdoc, testPath_in, null);

                // save the result
                pdfdoc.Save(testPath_out, SDFDoc.SaveOptions.e_linearized);

                // And we're done!
                // Console.WriteLine("Saved " + output_filename);
            }












            //{
            //    // If using the Professional version, put your serial key below.
            //    ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            //    // Continue to use the component in Trial mode when Free limit is reached.
            //    ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            //    // Load a PowerPoint file into the PresentationDocument object.
            //    var presentation = PresentationDocument.Load(testPath_in);

            //    // Create image save options.
            //    var imageOptions = new  GemBox.Presentation.ImageSaveOptions(ImageSaveFormat.Png)
            //    {
            //        SlideNumber = 0, // Select the first slide.
            //        Width = 1240 // Set the image width and keep the aspect ratio.
            //    };

            //    // Save the PresentationDocument object to a PNG file.
            //    presentation.Save(testPath_out, imageOptions);

            //}

            //{
            //    //Opens a PowerPoint Presentation
            //    Syncfusion.Presentation.IPresentation presentation = Syncfusion.Presentation.Presentation.Open(testPath_in);
            //    //Converts the PowerPoint Presentation into PDF document
            //    Syncfusion.Pdf.PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);
            //    //Saves the PDF document
            //    pdfDocument.Save(testPath_out);
            //    //Closes the PDF document
            //    pdfDocument.Close(true);
            //    //Closes the Presentation
            //    presentation.Close();
            //    //This will open the PDF file so, the result will be seen in default PDF viewer
            //    // System.Diagnostics.Process.Start("PPTToPDF.pdf");
            //    msg = "OK";
            //}












            //{

            //   Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            //    Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation = pptApplication.Presentations
            //    .Open(testPath_in, MsoTriState.msoFalse, MsoTriState.msoFalse
            //    , MsoTriState.msoFalse);

            //    pptPresentation.SaveAs(testPath_out, PpSaveAsFileType.ppSaveAsPDF);
            //    pptPresentation.Close();
            //    msg = "OK";
            //}









            //{
            //    //We assume that GdPicture has been correctly installed and unlocked.
            //    using (GdPictureDocumentConverter oConverter = new GdPictureDocumentConverter())
            //    {
            //        LicenseManager licenseManager = new LicenseManager();
            //        licenseManager.RegisterKEY("LICENSE_KEY");
            //        //Select your source document and its document format.
            //        GdPictureStatus status = oConverter.LoadFromFile(testPath_in, GdPicture14.DocumentFormat.DocumentFormatPPTX);
            //        if (status == GdPictureStatus.OK)
            //        {

            //            //Select the conformance of the resulting PDF document.
            //            status = oConverter.SaveAsPDF(testPath_out, PdfConformance.PDF);
            //            if (status == GdPictureStatus.OK)
            //            {
            //                msg = "OK";
            //            }
            //            else
            //            {
            //                msg = "NG";

            //            }
            //        }
            //        else
            //        {
            //            msg = "load fail";
            //        }
            //    }
            //}









            //{
            //    //Initialize an instance of the Presentation class
            //    //Initialize an instance of the Presentation class
            //    Spire.Presentation.Presentation ppt = new Spire.Presentation.Presentation();
            //    //Load a PowerPoint presentation
            //    ppt.LoadFromFile(testPath_in);

            //    //Specify the file path of the output HTML file 


            //    //Save the PowerPoint presentation to HTML format
            //    ppt.SaveToFile(testPath_out, Spire.Presentation.FileFormat.Html);

            //}






            //using (Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(testPath_in))
            //{
            //    HtmlOptions htmlOpt = new HtmlOptions();

            //    INotesCommentsLayoutingOptions options = htmlOpt.NotesCommentsLayouting;
            //    options.NotesPosition = NotesPositions.BottomFull;

            //    htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

            //    // Saves the presentation to HTML
            //    presentation.Save(testPath_out, Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
            //}

            return testPath_in;
        }


        public void deleteFileStoreInTempFolder(string fileFullPath)
        {
            // Xóa file trong Temp Folder
            try
            {

                System.IO.File.Delete(fileFullPath);

            }
            catch
            {

            }
        }

        public string AddPageNoForPDF(string fullPath, string outPath)
        {
            string oldFile = fullPath;

            string fileName = Path.GetFileName(oldFile);
            fileName = fileName.Replace("[Merge]", "");
            string newFilePath = outPath + "[PDF]" + fileName;
            bool status = false;
            try
            {
                // open the reader
                PdfReader reader = new PdfReader(oldFile);
                iTextSharp.text.Document document = new iTextSharp.text.Document();

                // open the writer
                FileStream fs = new FileStream(newFilePath, FileMode.Create, FileAccess.Write);
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

                    var multiLineString = "AUTO_PAGE: (" + System.Convert.ToString(i) + "/" + System.Convert.ToString(reader.NumberOfPages) + ")";
                    //var multiLineString = "";
                    contentByte.AddTemplate(importedPage, 0, 0);

                    contentByte.ShowTextAligned(PdfContentByte.ALIGN_LEFT, multiLineString, importedPage.Width - 100, 15, 0);
                    contentByte.EndText();

                }
                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();
                status = true;

            }
            catch
            {
                status = false;
            }
            return newFilePath;
        }

        public string CombineMultiplePDFs(string[] fileNames, string outFile)
        {
            // step 1: creation of a document-object
            iTextSharp.text.Document document = new iTextSharp.text.Document();
            //create newFileStream object which will be disposed at the end
            using (FileStream newFileStream = new FileStream(outFile, FileMode.Create))
            {
                // step 2: we create a writer that listens to the document
                PdfCopy writer = new PdfCopy(document, newFileStream);

                // step 3: we open the document
                document.Open();

                foreach (string fileName in fileNames)
                {
                    // we create a reader for a certain document
                    PdfReader reader = new PdfReader(fileName);
                    reader.ConsolidateNamedDestinations();

                    // step 4: we add content
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }

                    PRAcroForm form = reader.AcroForm;
                    //if (form != null)
                    //{
                    //    writer.CopyAcroForm(reader);
                    //    copy.setMergeFields();
                    //}

                    reader.Close();
                }

                // step 5: we close the document and writer
                writer.Close();
                document.Close();

            }//disposes the newFileStream object
            return outFile;
        }

        public void deleteWorksheetLisence(string fullPath)
        {
            //*use epluse
            try
            {
                using (ExcelPackage pck = new ExcelPackage(fullPath))
                {

                    var wk = pck.Workbook.Worksheets["Evaluation Warning"];
                    pck.Workbook.Worksheets.Delete(wk);
                    pck.SaveAs(fullPath);
                }
            }
            catch
            {

            }



        }

        public string editPDFtoRemoveLisenceText(string fullPath)
        {

            string oldFile = fullPath;
            string currentPath = Path.GetDirectoryName(fullPath);
            string fileName = Path.GetFileName(oldFile);
            string newFilePath = currentPath + "\\" + "[REMOVE_LICSENE_TEXT]" + fileName;
            bool status = false;
            // open the reader
            PdfReader reader = new PdfReader(oldFile);
            iTextSharp.text.Document document = new iTextSharp.text.Document();

            // open the writer
            FileStream fs = new FileStream(newFilePath, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();
            try
            {
                // create the new page and add it to the pdf
                for (var i = 1; i <= reader.NumberOfPages; i++)
                {
                    var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    var importedPage = writer.GetImportedPage(reader, i);
                    document.SetPageSize(importedPage.BoundingBox);
                    document.NewPage();
                    var contentByte = writer.DirectContent;

                    contentByte.AddTemplate(importedPage, 0, 0);

                    float llx = 0;
                    float lly = document.Top + 20;
                    float urx = document.PageSize.Right;
                    float ury = document.TopMargin + document.Top;
                    Rectangle rec = new Rectangle(llx, lly, urx, ury)
                    {
                        BackgroundColor = BaseColor.WHITE
                    };
                    contentByte.Rectangle(rec);

                }
                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();
                status = true;
                deleteFileStoreInTempFolder(oldFile);
            }
            catch
            {
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();
                status = true;
                deleteFileStoreInTempFolder(oldFile);
            }
            return newFilePath;
        }

        public int tempGenNo()
        {
            var all = db.DSUMs.Where(x => x.SubmitDate.Value.Year == DateTime.Now.Year).OrderBy(x => x.DFMID).ToList();
            int i = 1;
            foreach (var item in all)
            {
                item.DSUMNo = "DFM" + item.SubmitDate.Value.ToString("yyMMdd-") + i + "-00";
                db.Entry(item).State = EntityState.Modified;
                db.SaveChanges();
                i++;
            }
            return i;
        }


        public string deteteWorkSheetNoRelatedBeforeHandal(string fullPath, string processType)
        {
            fullPath = Server.MapPath("/File/Attachment/Temp/test.xlsm");
            processType = "MO";
            string[] removeWorkSheet = new string[6] { "Evaluation Warning", "Att-01_ver4__(draft_1)_(Mo_die)", "Att-02_ver_4__(draft_)(_PX_die)", "Att-01_ver3_", "Att-02_ver_3_", "Hystory" };

            try
            {
                using (ExcelPackage pck = new ExcelPackage(fullPath))
                {

                    //foreach (var name in removeWorkSheet)
                    //{
                    //    var wk = pck.Workbook.Worksheets[name];
                    //    if(wk != null)
                    //    {
                    //        pck.Workbook.Worksheets.Delete(wk);
                    //    }
                    //}
                    var wb = pck.Workbook;
                    foreach (var wsheet in wb.Worksheets)
                    {
                        if (processType == "MO")
                        {
                            if (!wsheet.Name.Contains("Att-01_ver7__Mo"))
                            {
                                pck.Workbook.Worksheets.Delete(wsheet);
                            }
                        }
                        else
                        {
                            if (!wsheet.Name.Contains("Att-02_ver8__PX"))
                            {
                                pck.Workbook.Worksheets.Delete(wsheet);
                            }
                        }

                    }

                    pck.SaveAs(fullPath);
                }
            }
            catch
            {

            }
            return fullPath;
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
