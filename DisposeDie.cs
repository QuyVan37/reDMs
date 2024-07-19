using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.RightsManagement;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Windows.Interop;
using Aspose.Slides;
using DMS03.Models;
using Microsoft.Ajax.Utilities;
using Microsoft.Office.Core;
using OfficeOpenXml;
using PagedList;
using pdftron.PDF;
using pdftron.SDF;
using pdftron;
using Spire.Doc;
using ceTe.DynamicPDF.ReportWriter.Data;
using Aspose.Slides.Export.Web;
using Avalonia.Controls;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Net.PeerToPeer;
using Microsoft.Office.Interop.Excel;
using Aspose.Cells.Drawing;
using System.Security.Policy;
using static iTextSharp.text.pdf.AcroFields;
using System.Net.Http;
using System.Web.Script.Serialization;
using PagedList;
using System.Collections;
using System.Dynamic;

namespace DMS03.Controllers
{
    public class VerifyDisposeDiesController : Controller
    {
        private DMSEntities db = new DMSEntities();
        SendEmailController sendEmailJob = new SendEmailController();
        CommonFunctionController commoneFunc = new CommonFunctionController();
        public StoreProcudure storeProcudure = new StoreProcudure();
        private readonly HttpClient _httpClient;
        public VerifyDisposeDiesController()
        {
            _httpClient = new HttpClient();
            // Set up any necessary configurations for your HTTP client here.
        }
        // GET: VerifyDisposeDies
        public ActionResult Index()
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            ViewBag.QuaterID = new SelectList(db.Dispose_Quater.OrderByDescending(x => x.QuaterID)
                                    .Select(q => new
                                    {
                                        QuaterID = q.QuaterID,
                                        Quater = q.Quater + "-" + q.Status

                                    }), "QuaterID", "Quater");
            return View();
        }

        public ActionResult HisUploadFile()
        {
            int currentYear = DateTime.Now.Year;
            var files = db.Dispose_ControlFileUpload.Where(x => x.UploadDate.Value.Year > currentYear - 3).OrderByDescending(x => x.FileID).ToList();
            return View(files);
        }
        //public JsonResult getData(string partNo, string assetNo, string waitFor, int? quaterID, string w_for)
        //{
        //    object data = null;
        //    if (!String.IsNullOrEmpty(w_for))
        //    {
        //        bool W_PE1 = w_for == "W_PE1" ? true : false;
        //        bool W_PE2 = w_for == "W_PE2" ? true : false;
        //        bool W_PAE = w_for == "W_PAE" ? true : false;
        //        bool W_PDC = w_for == "W_PDC" ? true : false;
        //        bool W_PUC = w_for == "W_PUC" ? true : false;
        //        bool W_DMT = w_for == "W_DMT" ? true : false;
        //        bool W_CRG = w_for == "W_CRG" ? true : false;
        //        bool W_CAM = w_for == "W_CAM" ? true : false;
        //        bool W_PUR = w_for == "W_PUR" ? true : false;
        //        data = storeProcudure.GetListVerifyDisposeByDept(W_PE1, W_PE2, W_PAE, W_DMT, W_PUC, W_CRG, W_PDC, W_CAM, W_PUR);

        //    }
        //    else
        //    {
        //        data = storeProcudure.getListVerifyForDisposeByParameters(partNo, assetNo, quaterID);
        //    }

        //    return Json(data, JsonRequestBehavior.AllowGet);
        //}

       
        public async Task<JsonResult> getData(string partNo, string assetNo, string waitFor, int? quaterID, string w_for, int? page)
        {
            if (page == null) page = 1;
            int pageSize = 2000;
            int pageNumber = (page ?? 1);


            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            if (!String.IsNullOrEmpty(w_for))
            {
                bool W_PE1 = w_for == "W_PE1" ? true : false;
                bool W_PE2 = w_for == "W_PE2" ? true : false;
                bool W_PAE = w_for == "W_PAE" ? true : false;
                bool W_PDC = w_for == "W_PDC" ? true : false;
                bool W_PUC = w_for == "W_PUC" ? true : false;
                bool W_DMT = w_for == "W_DMT" ? true : false;
                bool W_CRG = w_for == "W_CRG" ? true : false;
                bool W_CAM = w_for == "W_CAM" ? true : false;
                bool W_PUR = w_for == "W_PUR" ? true : false;
                // Thuc te ko su dung store nay vi phuc tap va ko chinh xac
                data = await storeProcudure.GetListVerifyDisposeByDeptAsyn(W_PE1, W_PE2, W_PAE, W_DMT, W_PUC, W_CRG, W_PDC, W_CAM, W_PUR);

            }
            else
            {
                data = await storeProcudure.GetListVerifyForDisposeByParametersAsync(partNo, assetNo, quaterID);
            }

            double totalPage = decimal.ToDouble(data.Count()) / decimal.ToDouble(pageSize);
            data = data.ToPagedList(pageNumber, pageSize).ToList();
            var output = new
            {
                test = totalPage,
                page = pageNumber,
                totalPage = Math.Ceiling(totalPage),
                data = data,
            };


            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getVerifyItemDetail(int verifyID)
        {

            var output = db.VerifyDisposeDies.Find(verifyID);
            var data = new
            {
                VerifyID = output.VerifyID,
                GenaralInfor = output.GenaralInfor,
                DieNo = output.DieNo
            };
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult saveData(int id, string name, string value)
        {

            bool status = false;
            string msg = "Update Fail";
            var today = DateTime.Now;
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            string userName = Session["Name"]?.ToString();
            string dept = Session["Dept"]?.ToString();
            string dispose_Role = Session["Dispose_Role"]?.ToString();
            var item = db.VerifyDisposeDies.Find(id);
            var supplierCode = db.InventoryResults.Find(item.InventoryID).LocationSupplierCode?.ToUpper();
            bool isInhouseDie = (supplierCode.Contains("MO") || supplierCode.Contains("MS") || supplierCode.Contains("DMT") || supplierCode.Contains("5400") || supplierCode.Contains("5500") || supplierCode.Contains("HOUSE"));
            if (item.Dispose_Quater.Status == "CLOSE")
            {
                status = false;
                msg = "Can not update for item already CLOSE";
                goto Exit;
            }

            // all Dept
            if (dispose_Role == "Check")
            {
                if (name == "GenaralInfor")
                {
                    item.GenaralInfor = today.ToString("yyyy/MM/dd_") + userName + ": " + value + System.Environment.NewLine + item.GenaralInfor;
                    // Luu log
                    status = true;
                    goto Save;
                }

                if (item.DieBelong != "LBP" && item.DieBelong != "CRG")
                {
                    if (name == "Manual_DieStatus")
                    {
                        item.Manual_DieStatus = value;
                        // Luu log
                        status = true;
                        goto Save;
                    }

                    if (name == "Manual_DieRepairHistory")
                    {
                        item.Manual_DieRepairHistory = value;
                        // Luu log
                        status = true;
                        goto Save;
                    }

                    if (name == "Manual_QtyOfDie")
                    {
                        int num = 0;
                        bool isCverOK = int.TryParse(value, out num);
                        if (isCverOK)
                        {
                            item.Manual_QtyOfDie = num;
                        }
                        else
                        {
                            if (value == "")
                            {
                                item.Manual_QtyOfDie = null;
                            }
                            else
                            {
                                msg = "Total die in use phải là số";
                                goto Exit;
                            }


                        }

                        status = true;
                        goto Save;
                    }
                    if (name == "Manual_Capacity")
                    {
                        double num = 0;
                        bool isCverOK = double.TryParse(value, out num);

                        if (isCverOK)
                        {
                            item.Manual_Capacity = num;
                        }
                        else
                        {
                            if (value == "")
                            {
                                item.Manual_Capacity = null;
                            }
                            else
                            {
                                msg = "Capacity phải là số";
                                goto Exit;
                            }
                        }

                        item.PURUpdateBy = userName;
                        item.PURUpdateDate = today;
                        status = true;
                    }
                }

            }




            // CAM
            if ((dept == "CAM" && dispose_Role == "Check") || admin == "Admin")
            {
                if (name == "DieBelong")
                {
                    item.DieBelong = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }

                // CAM verified từ kết quả các phòng ban input

                if (name == "CAM_KeepOrDispose")
                {
                    if (String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "Can not just when others dept not finish confirm";
                        goto Exit;
                    }
                    item.CAM_KeepOrDispose = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "CAM_ReasonOfDecision")
                {
                    if (String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "Can not just when others dept not finish confirm";
                        goto Exit;
                    }
                    item.CAM_ReasonOfDecision = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "CAM_NextVerifyStep")
                {
                    item.CAM_NextVerifyStep = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "CAM_DeptInCharge")
                {
                    item.CAM_DeptInCharge = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "CAM_Remak")
                {
                    item.CAM_Remak = value;
                    // Luu log
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                // Decision
                if (name == "Pre_Decision")
                {
                    if (String.IsNullOrEmpty(item.PUC_DMT_ConfirmKeepOrDispose))
                    {
                        msg = "DMT/PUC  chưa xác nhận kết quả servey với supplier KEEP or DISPOSE. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    item.PreDecisionID = String.IsNullOrEmpty(value) ? new Nullable<Int32>() : value.Contains("KEEP") ? 1 : 2;
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "Final_Decision")
                {
                    if (item.PreDecisionID == null)
                    {
                        msg = "Bạn chưa xác nhận kết quả Pre-Decision. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    item.FinalDecisionID = String.IsNullOrEmpty(value) ? new Nullable<Int32>() : value.Contains("KEEP") ? 1 : 2;
                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "TopApproveDate")
                {
                    if (item.FinalDecisionID == null)
                    {
                        msg = "Bạn chưa xác nhận kết quả Final Decision là KEEP or DISPOSE. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    DateTime dateConvert = DateTime.Now;
                    bool isDate = DateTime.TryParse(value, out dateConvert);
                    if (isDate == true || value == "")
                    {
                        item.TopApproveDate = isDate == true ? dateConvert : new Nullable<DateTime>();
                    }
                    else
                    {
                        msg = "Phải định dạng ngày (date)";
                        goto Exit;
                    }

                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }
                if (name == "PhysicalDisposeDate")
                {
                    if (item.TopApproveDate == null)
                    {
                        msg = "Bạn chưa xác nhận Top Approve Date. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    DateTime dateConvert = DateTime.Now;
                    bool isDate = DateTime.TryParse(value, out dateConvert);
                    if (isDate == true || value == "")
                    {
                        item.PhysicalDisposeDate = isDate == true ? dateConvert : new Nullable<DateTime>();
                    }
                    else
                    {
                        msg = "Phải định dạng ngày (date)";
                        goto Exit;
                    }

                    item.CAMUpdateBy = userName;
                    item.CAMUPdateDate = today;
                    status = true;
                }

            }

            // PUR/ DMT && name = PUR_TotalCav || name == "PUR_CycleTime"
            if ((((dept == "PUR" || dept == "DMT") && dispose_Role == "Check") || admin == "Admin") && (name == "PUR_TotalCav" || name == "PUR_CycleTime"))
            {
                if (!(admin == "Admin"))
                {
                    if (dept == "DMT" && !isInhouseDie)
                    {
                        msg = "Item này không thuộc INHOSE, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "PUR" && isInhouseDie)
                    {
                        msg = "Item này không thuộc OUTSIDE Supplier, Bạn không thể confirm";
                        goto Exit;
                    }
                }

                if (name == "PUR_TotalCav")
                {
                    int num = 0;
                    bool isCverOK = int.TryParse(value, out num);
                    if (isCverOK)
                    {
                        item.PUR_TotalCav = num;
                    }
                    else
                    {
                        if (value == "")
                        {
                            item.PUR_TotalCav = null;
                        }
                        else
                        {
                            msg = "Total Cav phải là số";
                            goto Exit;
                        }


                    }

                    item.PURUpdateBy = userName;
                    item.PURUpdateDate = today;
                    status = true;

                }
                if (name == "PUR_CycleTime")
                {
                    double num = 0;
                    bool isCverOK = double.TryParse(value, out num);

                    if (isCverOK)
                    {
                        item.PUR_CycleTime = num;
                    }
                    else
                    {
                        if (value == "")
                        {
                            item.PUR_CycleTime = null;
                        }
                        else
                        {
                            msg = "Cycle time phải là số";
                            goto Exit;
                        }
                    }

                    item.PURUpdateBy = userName;
                    item.PURUpdateDate = today;
                    status = true;
                }

            }



            //PAE/ DMT/ CRG/ PUC && name = PAE_CRG_DMT_PUC_Remark
            var allowDept = "PAE,DMT,CRG,PUC";
            if (((allowDept.Contains(dept) && dispose_Role == "Check") || admin == "Admin") && name == "PAE_CRG_DMT_PUC_Remark")
            {
                if (!(admin == "Admin"))
                {
                    if (dept == "DMT" && !isInhouseDie)
                    {
                        msg = "Item này không thuộc INHOSE, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "CRG" && !(item.DieBelong.Contains("CRG")))
                    {
                        msg = "Item này không thuộc CRG, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "PAE" && (!(item.DieBelong.Contains("LBP")) || isInhouseDie))
                    {
                        msg = "Item này không thuộc PAE response, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "PUC" && !(item.DieBelong.Contains("PACK") || item.DieBelong.Contains("ELEC") || item.DieBelong.Contains("OTHER")))
                    {
                        msg = "Item này không thuộc PUC response, Bạn không thể confirm";
                        goto Exit;
                    }
                }

                if (name == "PAE_CRG_DMT_PUC_Remark")
                {
                    item.PAE_CRG_DMT_PUC_Remark = value;

                    item.PAEUpdateBy = userName;
                    item.PAEUpdateDate = today;
                    status = true;
                }

            }




            // PE1/PE2/CRG && 

            if ((((dept.Contains("PE") || dept == "CRG") && dispose_Role == "Check") || admin == "Admin") && (name == "PE_AllModel" || name == "PE_CommonPart_CurrentModel" || name == "PE_CommonPart_NewModel" || name == "PE_AlternativePart" || name == "PE_FamilyPart"))
            {
                if (!(admin == "Admin"))
                {
                    if (dept == "PE1" && (item.DieBelong.Contains("ELEC") || item.DieBelong.Contains("CRG")))
                    {
                        msg = "Item này không thuộc PE1 responsibility, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "CRG" && !item.DieBelong.Contains("CRG"))
                    {
                        msg = "Item này không thuộc CRG responsibility, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "PE2" && !item.DieBelong.Contains("ELEC"))
                    {
                        msg = "Item này không thuộc PE2 responsibility, Bạn không thể confirm";
                        goto Exit;
                    }
                }


                if (name == "PE_AllModel")
                {
                    if (!String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "PDC đã sử dụng dữ liệu của PE để kiểm tra tình trạng demand. Vì vậy bạn ko thể sửa thông tin này. Bạn cần thông tin cho PDC xóa confirm của PDC trước, sau đó bạn mới có thể sửa thông tin này! ";
                        goto Exit;
                    }
                    item.PE_AllModel = value;
                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PE_AllModel = value;
                        if (dept == "PE1" || admin == "Admin")
                        {
                            c.PE1UpdateBy = userName;
                            c.PE1UpdateDate = today;
                        }
                        if (dept == "PE2" || admin == "Admin")
                        {
                            c.PE2UpdateBy = userName;
                            c.PE2UpdateDate = today;
                        }
                        if (dept == "CRG" || admin == "Admin")
                        {
                            c.CRGUpdateBy = userName;
                            c.CRGUpdateDate = today;
                        }
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;

                }
                if (name == "PE_CommonPart_CurrentModel")
                {
                    if (!String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "PDC đã sử dụng dữ liệu của PE để kiểm tra tình trạng demand. Vì vậy bạn ko thể sửa thông tin này. Bạn cần thông tin cho PDC xóa confirm của PDC trước, sau đó bạn mới có thể sửa thông tin này! ";
                        goto Exit;
                    }
                    item.PE_CommonPart_CurrentModel = value;
                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PE_CommonPart_CurrentModel = value;
                        if (dept == "PE1" || admin == "Admin")
                        {
                            c.PE1UpdateBy = userName;
                            c.PE1UpdateDate = today;
                        }
                        if (dept == "PE2" || admin == "Admin")
                        {
                            c.PE2UpdateBy = userName;
                            c.PE2UpdateDate = today;
                        }
                        if (dept == "CRG" || admin == "Admin")
                        {
                            c.CRGUpdateBy = userName;
                            c.CRGUpdateDate = today;
                        }
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;
                }

                if (name == "PE_CommonPart_NewModel")
                {
                    if (!String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "PDC đã sử dụng dữ liệu của PE để kiểm tra tình trạng demand. Vì vậy bạn ko thể sửa thông tin này. Bạn cần thông tin cho PDC xóa confirm của PDC trước, sau đó bạn mới có thể sửa thông tin này! ";
                        goto Exit;
                    }
                    item.PE_CommonPart_NewModel = value;
                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PE_CommonPart_NewModel = value;
                        if (dept == "PE1" || admin == "Admin")
                        {
                            c.PE1UpdateBy = userName;
                            c.PE1UpdateDate = today;
                        }
                        if (dept == "PE2" || admin == "Admin")
                        {
                            c.PE2UpdateBy = userName;
                            c.PE2UpdateDate = today;
                        }
                        if (dept == "CRG" || admin == "Admin")
                        {
                            c.CRGUpdateBy = userName;
                            c.CRGUpdateDate = today;
                        }
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;
                }

                if (name == "PE_AlternativePart")
                {
                    if (!String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "PDC đã sử dụng dữ liệu của PE để kiểm tra tình trạng demand. Vì vậy bạn ko thể sửa thông tin này. Bạn cần thông tin cho PDC xóa confirm của PDC trước, sau đó bạn mới có thể sửa thông tin này! ";
                        goto Exit;
                    }
                    item.PE_AlternativePart = value;

                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PE_AlternativePart = value;
                        if (dept == "PE1" || admin == "Admin")
                        {
                            c.PE1UpdateBy = userName;
                            c.PE1UpdateDate = today;
                        }
                        if (dept == "PE2" || admin == "Admin")
                        {
                            c.PE2UpdateBy = userName;
                            c.PE2UpdateDate = today;
                        }
                        if (dept == "CRG" || admin == "Admin")
                        {
                            c.CRGUpdateBy = userName;
                            c.CRGUpdateDate = today;
                        }
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;
                }
                if (name == "PE_FamilyPart")
                {
                    if (!String.IsNullOrEmpty(item.PDC_PartDemandStatus))
                    {
                        msg = "PDC đã sử dụng dữ liệu của PE để kiểm tra tình trạng demand. Vì vậy bạn ko thể sửa thông tin này. Bạn cần thông tin cho PDC xóa confirm của PDC trước, sau đó bạn mới có thể sửa thông tin này! ";
                        goto Exit;
                    }
                    item.PE_FamilyPart = value;

                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PE_FamilyPart = value;
                        if (dept == "PE1" || admin == "Admin")
                        {
                            c.PE1UpdateBy = userName;
                            c.PE1UpdateDate = today;
                        }
                        if (dept == "PE2" || admin == "Admin")
                        {
                            c.PE2UpdateBy = userName;
                            c.PE2UpdateDate = today;
                        }
                        if (dept == "CRG" || admin == "Admin")
                        {
                            c.CRGUpdateBy = userName;
                            c.CRGUpdateDate = today;
                        }
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    status = true;
                }

            }

            // PDC
            if ((dept == "PDC" && dispose_Role == "Check") || admin == "Admin")
            {
                if (name == "PDC_PartDemandStatus")
                {
                    if (String.IsNullOrEmpty(item.PE_AllModel) || String.IsNullOrEmpty(item.PE_AlternativePart) || String.IsNullOrEmpty(item.PE_CommonPart_CurrentModel) || String.IsNullOrEmpty(item.PE_CommonPart_NewModel) || String.IsNullOrEmpty(item.PE_FamilyPart) && !String.IsNullOrEmpty(value))
                    {
                        msg = "PE chưa xác nhận đầy đủ thông tin model, common và family part. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    item.PDC_PartDemandStatus = value;
                    item.PDC1UpdateBy = userName;
                    item.PDC1UpdateDate = today;
                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PDC_PartDemandStatus = value;
                        c.PDC1UpdateBy = userName;
                        c.PDC1UpdateDate = today;
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    status = true;
                }
                if (name == "PDC_MP_MaxDemand")
                {
                    if (String.IsNullOrEmpty(item.PE_AllModel) || String.IsNullOrEmpty(item.PE_AlternativePart) || String.IsNullOrEmpty(item.PE_CommonPart_CurrentModel) || String.IsNullOrEmpty(item.PE_CommonPart_NewModel) || String.IsNullOrEmpty(item.PE_FamilyPart))
                    {
                        msg = "PE chưa xác nhận đầy đủ thông tin model, common và family part. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    double num = 0;
                    bool isCVOK = Double.TryParse(value, out num);
                    if (isCVOK)
                    {
                        item.PDC_MP_MaxDemand = num;
                    }
                    else
                    {
                        item.PDC_MP_MaxDemand = null;
                        msg = "Max Demand shot phải là số";
                        goto Exit;
                    }
                    item.PDC1UpdateBy = userName;
                    item.PDC1UpdateDate = today;

                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        if (isCVOK)
                        {
                            c.PDC_MP_MaxDemand = num;
                        }
                        else
                        {
                            c.PDC_MP_MaxDemand = null;
                        }
                        c.PDC1UpdateBy = userName;
                        c.PDC1UpdateDate = today;
                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    status = true;
                }
                if (name == "PDC_ResultJP_FB")
                {
                    item.PDC_ResultJP_FB = value;
                    item.PDC1UpdateBy = userName;
                    item.PDC1UpdateDate = today;

                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PDC_ResultJP_FB = value;
                        c.PDC1UpdateBy = userName;
                        c.PDC1UpdateDate = today;

                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;
                }
                if (name == "PDC_Remark")
                {
                    item.PDC_Remark = value;
                    item.PDC1UpdateBy = userName;
                    item.PDC1UpdateDate = today;

                    //Code Update same PartNo
                    var sameItem = db.VerifyDisposeDies.Where(x => x.DieNo.Contains(item.DieNo.Substring(0, 8)) && x.Dispose_Quater.Status == "OPEN").ToList();
                    foreach (var c in sameItem)
                    {
                        c.PDC_ResultJP_FB = value;
                        c.PDC1UpdateBy = userName;
                        c.PDC1UpdateDate = today;

                        db.Entry(c).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    status = true;
                }
            }


            // PUC/DMT Servey
            var allowDept1 = "DMT,PUC";
            if (((allowDept1.Contains(dept) && dispose_Role == "Check") || admin == "Admin") && (name == "PUC_WarrantyShot" || name == "PUC_DMT_ConfirmKeepOrDispose" || name == "PUC_DMT_ReasonKeepOrDispose" || name == "PUC_DMT_ConfirmHasCommonOrFamily" || name == "PUC_DMT_Reason_NotMatch"))
            {
                if (!(admin == "Admin"))
                {
                    if (dept == "DMT" && !isInhouseDie)
                    {
                        msg = "Item này không thuộc INHOUSE, Bạn không thể confirm";
                        goto Exit;
                    }
                    if (dept == "PUC" && isInhouseDie)
                    {
                        msg = "Item này không thuộc OUTSIDE Supplier, Bạn không thể confirm";
                        goto Exit;
                    }
                }

                if (name == "PUC_WarrantyShot")
                {
                    int num = 0;
                    bool isCVOK = int.TryParse(value, out num);
                    if (isCVOK)
                    {
                        item.PUC_WarrantyShot = num;
                    }
                    else
                    {
                        if (value == "")
                        {
                            item.PUC_WarrantyShot = null;
                        }
                        else
                        {
                            msg = "Garantee shot phải là số";
                            goto Exit;
                        }

                    }

                    item.PUCUpdateBy = userName;
                    item.PUCUpdateDate = today;
                    status = true;
                }
                if (name == "PUC_DMT_ConfirmKeepOrDispose")
                {
                    if (String.IsNullOrEmpty(item.CAM_KeepOrDispose))
                    {
                        msg = "CAM chưa xác nhận KEEP or DISPOSE. Vì vậy bạn chưa thể xác nhận thôn tin này!";
                        goto Exit;
                    }
                    item.PUC_DMT_ConfirmKeepOrDispose = value;
                    item.PUCUpdateBy = userName;
                    item.PUCUpdateDate = today;
                    status = true;
                }
                if (name == "PUC_DMT_ReasonKeepOrDispose")
                {
                    if (String.IsNullOrEmpty(item.CAM_KeepOrDispose))
                    {
                        msg = "CAM chưa xác nhận KEEP or DISPOSE. Vì vậy bạn chưa thể xác nhận thôn tin này!";
                        goto Exit;
                    }
                    item.PUC_DMT_ReasonKeepOrDispose = value;
                    item.PUCUpdateBy = userName;
                    item.PUCUpdateDate = today;
                    status = true;
                }
                if (name == "PUC_DMT_ConfirmHasCommonOrFamily")
                {
                    item.PUC_DMT_ConfirmHasCommonOrFamily = value;
                    item.PUCUpdateBy = userName;
                    item.PUCUpdateDate = today;
                    status = true;
                }
                if (name == "PUC_DMT_Reason_NotMatch")
                {
                    if (String.IsNullOrEmpty(item.CAM_KeepOrDispose))
                    {
                        msg = "CAM chưa xác nhận KEEP or DISPOSE. Vì vậy bạn chưa thể xác nhận thông tin này!";
                        goto Exit;
                    }
                    item.PUC_DMT_Reason_NotMatch = value;
                    item.PUCUpdateBy = userName;
                    item.PUCUpdateDate = today;
                    status = true;
                }


            }






        Save:
            //Luu log
            item.Log_HisUpdate = today.ToString("yyyyMMdd_HH:mm tt ") + userName + " updated [" + name + ": " + value + "]" + System.Environment.NewLine + item.Log_HisUpdate;
            db.Entry(item).State = EntityState.Modified;
            db.SaveChanges();


        Exit:
            var output = new
            {
                status = status,
                msg = msg
            };
            return Json(output, JsonRequestBehavior.AllowGet);


        }

        public JsonResult getListDieRefer(string partNo)
        {
            var data = storeProcudure.getListDieRefer(partNo);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CAMChangeDecision(int id, string reason, HttpPostedFileBase fileEvedent)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            bool status = false;
            var msg = "";
            if (String.IsNullOrWhiteSpace(reason))
            {
                msg = "You need input reason!";
                goto Exit;
            }

            if (dept == "CAM" && (role == "Check" || role == "Approve"))
            {
                var item = db.VerifyDisposeDies.Find(id);
                int? currentDecisionID = item.FinalDecisionID;
                if (currentDecisionID == 2) // Dispose
                {
                    item.FinalDecisionID = 1; // Keep
                    item.ReasonChangeDecision = reason;
                }

                if (fileEvedent != null)
                {
                    var fileExt = Path.GetExtension(fileEvedent.FileName);
                    // Code luu file vào folder
                    var fileName = "Evendent_ChangeDecision_" + item.DieNo + DateTime.Now.ToString("yyyy-MM-dd-HHmmss") + fileExt;
                    var path = Path.Combine(Server.MapPath("~/File/Attachment/"), fileName);
                    fileEvedent.SaveAs(path);
                    item.FileEvendentChangeDecision = fileName;
                }
                db.Entry(item).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                msg = "You do not have permission!";
            }
        Exit:
            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }


        public ActionResult CAMAnnoucement(int? fileID, string noneOPInvent, string noneOPDMS, string none_OP, string alreadyVerify, string matchDMS, string addedAndRefered, string added, string refered, string search, int? page, string export)
        {
            if (page == null) page = 1;
            int pageSize = 50;
            int pageNumber = (page ?? 1);

            var output = db.InventoryResults.Where(x => x.FileInventID == fileID && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();
            if (!String.IsNullOrEmpty(matchDMS) && fileID > 0)
            {
                output = db.InventoryResults.Where(x => x.FileInventID == fileID && x.IsMatchDMSDatabase == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();

            }

            if (!String.IsNullOrEmpty(alreadyVerify) && fileID > 0)
            {
                output = db.InventoryResults.Where(x => x.FileInventID == fileID && x.isVerifiedDisposeLastQuater == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();

            }

            if (!String.IsNullOrEmpty(noneOPDMS) && fileID > 0)
            {
                output = db.InventoryResults.Where(x => x.FileInventID == fileID && x.isNoneOperationFollowDMSCheck == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();

            }

            if (!String.IsNullOrEmpty(noneOPInvent) && fileID > 0)
            {
                output = db.InventoryResults.Where(x => x.FileInventID == fileID && x.isNoneOperationFollowInventory == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();
            }

            if (!String.IsNullOrEmpty(added))
            {
                output = db.InventoryResults.Where(x => x.IsSeletedForVerify == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();
            }
            if (!String.IsNullOrEmpty(refered))
            {
                output = db.InventoryResults.Where(x => x.isReferDie == true && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();
            }
            if (!String.IsNullOrEmpty(addedAndRefered))
            {
                output = db.InventoryResults.Where(x => (x.IsSeletedForVerify == true || x.isReferDie == true) && x.Active != false).Include(x => x.Dispose_ControlFileUpload).ToList();
            }

            if (!String.IsNullOrEmpty(none_OP) && fileID > 0)
            {
                output = db.InventoryResults.Where(x => x.FileInventID == fileID && (x.isNoneOperationFollowDMSCheck == true || x.isNoneOperationFollowInventory == true) && x.Active != false && (x.IsSeletedForVerify != true && x.isReferDie != true && x.isVerifiedDisposeLastQuater != true))
                                            .Include(x => x.Dispose_ControlFileUpload).ToList();
            }


            if (!String.IsNullOrEmpty(search))
            {
                search = search.Trim().ToUpper();
                output = db.InventoryResults.Where(x => x.AssetNo.Contains(search) && x.Active != false).OrderByDescending(x => x.Dispose_ControlFileUpload.UploadDate).ToList();

                if (output.Count == 0)
                {
                    output = db.InventoryResults.Where(x => x.DieNo.Contains(search) && x.Active != false).OrderByDescending(x => x.Dispose_ControlFileUpload.UploadDate).ToList();

                }

            }


            var fileUploadedInfor = db.Dispose_ControlFileUpload.Find(fileID);
            ViewBag.fileID = new SelectList(db.Dispose_ControlFileUpload.Where(x => x.Type.Contains("InventoryResult") && x.Active == true).OrderByDescending(x => x.UploadDate).Take(20).ToList(), "FileID", "FileName", fileID);
            ViewBag.FileName = fileUploadedInfor?.FileName;
            ViewBag.uploadedDate = fileUploadedInfor?.UploadDate;
            ViewBag.uploadedBy = fileUploadedInfor?.UploadBy;
            ViewBag.totalItem = fileUploadedInfor?.TotalItem;
            ViewBag.NoOfItemsMatchDMS = fileUploadedInfor?.NoOfItemsMatchDMS;
            ViewBag.NoOfItemNoneOperationfollowInvent = fileUploadedInfor?.NoOfItemNoneOperationfollowInvent;
            ViewBag.NoOfItemNoneOperationfollowDMS = fileUploadedInfor?.NoOfItemNoneOperationfollowDMS;
            ViewBag.NoOftemVerifiedDisposeLastQ = fileUploadedInfor?.NoOftemVerifiedDisposeLastQ;
            ViewBag.selectedItem = db.InventoryResults.Where(x => x.IsSeletedForVerify == true && x.Active != false).Count();
            ViewBag.totalRefer = db.InventoryResults.Where(x => x.isReferDie == true && x.Active != false).Count();
            ViewBag.search = search;
            ViewBag.FileIDSeleted = fileID;
            ViewBag.TotalView = output.Count();
            ViewBag.QuaterID = new SelectList(db.Dispose_Quater.Where(x => x.Status == "OPEN").OrderByDescending(x => x.QuaterID), "QuaterID", "Quater");
            if (!String.IsNullOrEmpty(export))
            {
                ExportListInventoryResult(output);
            }

            return View(output.ToPagedList(pageNumber, pageSize));
        }

        public JsonResult checkQuaterCanCloseOrNot(int quaterID)
        {
            bool status = false;
            var items = db.VerifyDisposeDies.Where(x => x.QuaterID == quaterID).ToList();
            var NoItemNYDecision = items.Where(x => x.FinalDecisionID == null).Count();
            var NoItemNYTopApprove = items.Where(x => x.TopApproveDate == null).Count();
            var NoOfItemDecidedDisposal = items.Where(x => x.FinalDecisionID == 2).Count();
            var NoOfItemNYPhysicalDispose = items.Where(x => x.FinalDecisionID == 2 && x.PhysicalDisposeDate == null).Count();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            var isPermition = false;
            if (admin == "Admin" || (dept == "CAM" && (role == "Check" || role == "Approve")))
            {
                isPermition = true;
            }
            if (NoOfItemNYPhysicalDispose == 0 && NoItemNYTopApprove == 0)
            {
                status = true;
            }

            return Json(new { status = status, isPermition = isPermition, NoItemNYDecision = NoItemNYDecision, NoItemNYTopApprove = NoItemNYTopApprove, NoOfItemNYPhysicalDispose = NoOfItemNYPhysicalDispose }, JsonRequestBehavior.AllowGet);

        }

        public JsonResult closeQuaterOpen(int quaterID)
        {
            bool status = false;
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            string msg = "";
            if (admin == "Admin" || (dept == "CAM" && (role == "Check" || role == "Approve")))
            {
                var items = db.VerifyDisposeDies.Where(x => x.QuaterID == quaterID).ToList();
                var NoItemNYDecision = items.Where(x => x.FinalDecisionID == null).Count();
                var NoItemNYTopApprove = items.Where(x => x.TopApproveDate == null).Count();
                var NoOfItemDecidedDisposal = items.Where(x => x.FinalDecisionID == 2).Count();
                var NoOfItemNYPhysicalDispose = items.Where(x => x.FinalDecisionID == 2 && x.PhysicalDisposeDate == null).Count();
                if (NoOfItemNYPhysicalDispose == 0 && NoItemNYTopApprove == 0)
                {
                    // chuyen các item ở inventory selected true => fasle
                    foreach (var item in items)
                    {
                        var invent = db.InventoryResults.Find(item.InventoryID);
                        invent.IsSeletedForVerify = false;
                        db.Entry(invent).State = EntityState.Modified;
                        db.SaveChanges();

                    }
                    // chuyển quater status from OPEN to CLOSE
                    var quater = db.Dispose_Quater.Find(quaterID);
                    quater.Status = "CLOSE";
                    db.Entry(quater).State = EntityState.Modified;
                    db.SaveChanges();

                    // Chuyen die status sang disposed
                    // Tạm thời deactive sẽ active sau khi offical
                    //storeProcudure.updateDieDatabaseAfterCloseVerifyQuater(quaterID);


                    // Update service part
                    storeProcudure.updateServicePartAfterCloseVerifyQuater(quaterID);




                    status = true;
                }
                else
                {
                    msg = "Still pending items NY decision: " + NoItemNYDecision + " and NY Top Approve: " + NoItemNYTopApprove;
                }

            }
            else
            {
                msg = "You do not have permission!!";
            }
            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult EditInventoryResult(int id)
        {
            // ONLY ADMIN CAN EDIT
            if (Session["Admin"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            if (admin != "Admin")
            {
                ViewBag.err = "Only Admin of Dispose Project (CAM) can edit";
            }
            var item = db.InventoryResults.Find(id);
            return View(item);
        }

        [HttpPost]
        public ActionResult EditInventoryResult(InventoryResult invent)
        {
            // ONLY ADMIN CAN EDIT
            if (Session["Admin"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            if (admin != "Admin")
            {
                ViewBag.err = "Only Admin of Dispose Project (CAM) can edit";
                return View(invent);
            }

            var ExistInvent = db.InventoryResults.Find(invent.InventoryID);

            ExistInvent.AssetNo = invent.AssetNo;
            ExistInvent.OldAssetNo = invent.OldAssetNo;
            ExistInvent.CostCenter = invent.CostCenter;
            ExistInvent.DeptName = invent.DeptName;
            ExistInvent.AssetName = invent.AssetName;
            ExistInvent.Note1 = invent.Note1;
            ExistInvent.ComponentAsset = invent.ComponentAsset;
            ExistInvent.ClassCode = invent.ClassCode;
            ExistInvent.OriginalCostUSD = invent.OriginalCostUSD;
            ExistInvent.RemainCostUSD = invent.RemainCostUSD;
            ExistInvent.StartUseDate = invent.StartUseDate;
            ExistInvent.LocationSupplierCode = invent.LocationSupplierCode;
            ExistInvent.DieNo = invent.DieNo;
            ExistInvent.IsFixAsset = invent.IsFixAsset;
            ExistInvent.FAPlate = invent.FAPlate;
            ExistInvent.UsingStatus = invent.UsingStatus;
            ExistInvent.StopDate = invent.StopDate;
            ExistInvent.ActionPlanForUnuse = invent.ActionPlanForUnuse;
            ExistInvent.ReasonforDispose = invent.ReasonforDispose;
            ExistInvent.Shot = invent.Shot;
            ExistInvent.RecordShotDate = invent.RecordShotDate;
            ExistInvent.IsWrongLocation = invent.IsWrongLocation;
            ExistInvent.ControlNumber = invent.ControlNumber;
            ExistInvent.ControlDept = invent.ControlDept;

            db.Entry(ExistInvent).State = EntityState.Modified;
            db.SaveChanges();
            // System checking
            // gọi store procedure update thông tin die và update thông tin inventory
            storeProcudure.procudureVerifyInventoryResult((int)ExistInvent.FileInventID);
            storeProcudure.procudureUpdateDieDatabaseFollowInventory((int)ExistInvent.FileInventID);
            ViewBag.suc = "Saved";
            return View(ExistInvent);
        }

        public JsonResult selectReferDie(string[] InventoryIDs, int? fileID)
        {
            InventoryIDs = InventoryIDs == null ? new string[0] { } : InventoryIDs.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            bool status = false;
            int success = 0;
            int fail = 0;
            string msg = "";
            if (InventoryIDs.Length > 0)
            {
                foreach (var idString in InventoryIDs)
                {
                    int id = int.Parse(idString);
                    var inventRecord = db.InventoryResults.Find(id);

                    if (inventRecord.isVerifiedDisposeLastQuater == true)
                    {
                        fail += 1;
                        msg += inventRecord.AssetNo + " already verify last Q" + Environment.NewLine;
                    }
                    else
                    {
                        inventRecord.isReferDie = true;
                        inventRecord.IsSeletedForVerify = false;
                        db.Entry(inventRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success += 1;

                    }

                }
                status = true;
            }

            // Total items verify
            var totalAdded = db.InventoryResults.Where(x => x.IsSeletedForVerify == true && x.Active != false).Count();
            var totalRefer = db.InventoryResults.Where(x => x.isReferDie == true && x.Active != false).Count();
            return Json(new { status = status, totalAdded = totalAdded, totalRefer = totalRefer, success = success, fail = fail, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult selectToVerify(string[] InventoryIDs, int? fileID, string NoneOPType)
        {
            InventoryIDs = InventoryIDs == null ? new string[0] { } : InventoryIDs.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            bool status = false;
            int success = 0;
            int fail = 0;
            string msg = "";
            if (InventoryIDs.Length > 0)
            {
                foreach (var idString in InventoryIDs)
                {
                    int id = int.Parse(idString);
                    var inventRecord = db.InventoryResults.Find(id);

                    if (inventRecord.isVerifiedDisposeLastQuater == true)
                    {
                        fail += 1;
                        msg += inventRecord.AssetNo + " already verify last Q" + Environment.NewLine;
                    }
                    else
                    {
                        inventRecord.IsSeletedForVerify = true;
                        db.Entry(inventRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success += 1;

                    }

                }
                status = true;
            }

            // Xu li input
            if (NoneOPType == "noneOPInvent")
            {
                var r = db.InventoryResults.Where(x => x.FileInventID == fileID && x.isNoneOperationFollowInventory == true && x.Active != false).ToList();
                foreach (var inventRecord in r)
                {
                    if (inventRecord.isVerifiedDisposeLastQuater == true)
                    {
                        fail += 1;
                        msg += inventRecord.AssetNo + " already verify last Q" + Environment.NewLine;
                    }
                    else
                    {
                        inventRecord.IsSeletedForVerify = true;
                        db.Entry(inventRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success += 1;

                    }
                }
                status = true;

            }

            // Xu li input
            if (NoneOPType == "noneOPDMS")
            {
                var r = db.InventoryResults.Where(x => x.FileInventID == fileID && x.isNoneOperationFollowDMSCheck == true && x.Active != false).ToList();
                foreach (var inventRecord in r)
                {
                    if (inventRecord.isVerifiedDisposeLastQuater == true)
                    {
                        fail += 1;
                        msg += inventRecord.AssetNo + " already verify last Q" + Environment.NewLine;

                    }
                    else
                    {
                        inventRecord.IsSeletedForVerify = true;
                        db.Entry(inventRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success += 1;
                    }
                }
                status = true;
            }

            // Total items verify
            var totalAdded = db.InventoryResults.Where(x => x.IsSeletedForVerify == true && x.Active != false).Count();
            var totalRefer = db.InventoryResults.Where(x => x.isReferDie == true && x.Active != false).Count();
            return Json(new { status = status, totalAdded = totalAdded, totalRefer = totalRefer, success = success, fail = fail, msg = msg }, JsonRequestBehavior.AllowGet);

        } // CAM Just select before announcement

        public JsonResult RemoveToVerify(string[] InventoryIDs)
        {
            InventoryIDs = InventoryIDs == null ? new string[0] { } : InventoryIDs.Where(x => !String.IsNullOrEmpty(x)).ToArray();
            bool status = false;
            int success = 0;
            int fail = 0;
            string msg = "";
            if (InventoryIDs.Length > 0)
            {
                foreach (var idString in InventoryIDs)
                {
                    int id = int.Parse(idString);
                    var inventRecord = db.InventoryResults.Find(id);

                    if (inventRecord.IsSeletedForVerify == true || inventRecord.isReferDie == true)
                    {
                        inventRecord.IsSeletedForVerify = false;
                        inventRecord.isReferDie = false;
                        db.Entry(inventRecord).State = EntityState.Modified;
                        db.SaveChanges();
                        success += 1;

                    }
                    else
                    {

                        fail += 1;
                        msg += inventRecord.AssetNo + " not in verify list." + Environment.NewLine;

                    }

                }
                status = true;
            }

            // Total items verify
            var totalAdded = db.InventoryResults.Where(x => x.IsSeletedForVerify == true && x.Active != false).Count();
            var totalRefer = db.InventoryResults.Where(x => x.isReferDie == true && x.Active != false).Count();
            return Json(new { status = status, totalAdded = totalAdded, totalRefer = totalRefer, success = success, fail = fail, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult PDCconfirmEOLAndServicePart()
        {
            if (Session["Dept"] == null)
            {
                Session["URL"] = HttpContext.Request.Url.PathAndQuery;
                return RedirectToAction("Index", "Login");
            }

            // ViewBag.EOLParts = db.Dispose_UniquePart.OrderByDescending(x => x.UniquePartID).Take(50).ToList();
            // ViewBag.ServiceParts = db.Dispose_ServicePart.OrderByDescending(x => x.ServicePartID).Take(50).ToList();
            return View();
        }

        public JsonResult PDCUploadListUniquePart(HttpPostedFileBase file)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            var name = Session["Name"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            bool status = false;
            string msg = "";
            var today = DateTime.Now;
            if (dept == "PDC" && (role == "Check" || role == "Approve") || admin == "Admin")
            {


                if (file == null)
                {
                    status = false;
                    msg = "Please select file Unique part EOL";
                    return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
                }
                // 1. Lưu file & control
                // save file vật lí
                string fileName = "Unique-Part-EOL-uploadDate-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/DisposeDie/");
                fileName += fileExt;
                file.SaveAs(path + fileName);


                Dispose_ControlFileUpload newFile = new Dispose_ControlFileUpload()
                {
                    FileName = fileName,
                    Type = "Unique_Part_EOL",
                    Dept = dept,
                    UploadBy = name,
                    UploadDate = today,
                    Active = true
                };
                db.Dispose_ControlFileUpload.Add(newFile);
                db.SaveChanges();


                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    // Check version
                    // 2. Kiểm tra format
                    var A3 = "Part No";
                    var B3 = "Model";
                    var C3 = "EOL Date";
                    var a3 = worksheet.Cells["A3"].Text.Trim();
                    var b3 = worksheet.Cells["B3"].Text.Trim();
                    var c3 = worksheet.Cells["C3"].Text.Trim();


                    if (A3 != a3 || B3 != b3 || C3 != c3)
                    {
                        status = false;
                        msg = "file: " + file.FileName + " is Wrong format!";
                        return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
                    }


                    // 3. Đọc dữ liệu và lưu Database
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    for (int row = start.Row + 3; row <= end.Row; row++)
                    { // Row by row...

                        try
                        {
                            string PartNo = worksheet.Cells[row, 1].Text.ToUpper().Trim();
                            string Model = worksheet.Cells[row, 2].Text.ToUpper().Trim();
                            string EOLDate = worksheet.Cells[row, 3].Text.ToUpper().Trim();

                            if (String.IsNullOrEmpty(PartNo)) break;

                            // lưu dữ liệu vào database
                            var checkExist = db.Dispose_UniquePart.Where(x => x.PartNo == PartNo).FirstOrDefault();
                            if (checkExist == null)
                            {
                                Dispose_UniquePart newUniquePart = new Dispose_UniquePart()
                                {
                                    PartNo = PartNo,
                                    Model = Model,
                                    EOLDate = EOLDate,
                                    UpdateBy = name,
                                    UpdateDate = today
                                };
                                db.Dispose_UniquePart.Add(newUniquePart);
                                db.SaveChanges();
                                checkExist = newUniquePart;
                            }
                            else
                            {
                                checkExist.Model = Model;
                                checkExist.EOLDate = EOLDate;
                                checkExist.UpdateBy = name;
                                checkExist.UpdateDate = today;
                                db.Entry(checkExist).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            status = true;

                            // Update Database
                            // check nếu die ko common hoặc family với part khác thì sẽ chuyển DieStatusID = 6 (EOL)
                            var checkExistPart = db.Parts1.Where(x => x.PartNo.Contains(PartNo) && x.Active != false).FirstOrDefault();
                            if (checkExistPart != null)
                            {
                                var ListCommon = db.CommonDie1.Where(x => x.PartID == checkExistPart.PartID && x.Active != false && (x.Die1.DieStatusID != 6 && x.Die1.DieStatusID != 8 && x.Die1.DieStatusID != 9 && x.Die1.DieStatusID != 10)).ToList();
                                bool isCommonOrFamily = false;
                                foreach (var common in ListCommon)
                                {
                                    // Neu chi can 1 comon ko chua partNO => co common Or family
                                    isCommonOrFamily = !common.DieNo.Contains(PartNo) ? true : false;
                                    if (isCommonOrFamily == true)
                                    {
                                        break;
                                    }
                                }
                                // neu ko common thi chuyen san status = 6
                                if (isCommonOrFamily == false)
                                {
                                    foreach (var common in ListCommon)
                                    {
                                        var die = db.Die1.Find(common.DieID);
                                        die.DieStatusID = 6;
                                        die.RemarkDieStatusUsing = "PDC confirmed Part Unique EOL";
                                        db.Entry(die).State = EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                }

                            }
                        }
                        catch
                        {

                        }

                    }

                }
            }

            else
            {
                msg = "You do not permission!";
            }

            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult PDCUploadListConfirmServicePart(HttpPostedFileBase file)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            var name = Session["Name"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            bool status = false;
            string msg = "";
            int suc = 0;
            int fail = 0;
            List<string> listFail = new List<string>();
            var today = DateTime.Now;
            if (dept == "PDC" && (role == "Check" || role == "Approve") || admin == "Admin")
            {


                if (file == null)
                {
                    status = false;
                    msg = "Please select file confirm service parts";
                    return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
                }
                // 1. Lưu file & control
                // save file vật lí
                string fileName = "Service-Part-Confirm-uploadDate-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/DisposeDie/");
                fileName += fileExt;
                file.SaveAs(path + fileName);


                Dispose_ControlFileUpload newFile = new Dispose_ControlFileUpload()
                {
                    FileName = fileName,
                    Type = "Service_Part_Confirm",
                    Dept = dept,
                    UploadBy = name,
                    UploadDate = today,
                    Active = true
                };
                db.Dispose_ControlFileUpload.Add(newFile);
                db.SaveChanges();


                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    // Check version
                    // 2. Kiểm tra format
                    var A3 = "Part No";
                    var B3 = "Model";
                    var C3 = "Lastest Confirm";
                    var D3 = "By";
                    var E3 = "Confrim Date";
                    var F3 = "New Confrim";
                    var a3 = worksheet.Cells["A3"].Text.Trim();
                    var b3 = worksheet.Cells["B3"].Text.Trim();
                    var c3 = worksheet.Cells["C3"].Text.Trim();
                    var d3 = worksheet.Cells["D3"].Text.Trim();
                    var e3 = worksheet.Cells["E3"].Text.Trim();
                    var f3 = worksheet.Cells["F3"].Text.Trim();


                    if (A3 != a3 || B3 != b3 || C3 != c3 || D3 != d3 || E3 != e3 || F3 != f3)
                    {
                        status = false;
                        msg = "file: " + file.FileName + " is Wrong format!";
                        return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
                    }


                    // 3. Đọc dữ liệu và lưu Database
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    for (int row = start.Row + 3; row <= end.Row; row++)
                    { // Row by row...
                        string temMsg = "";
                        try
                        {
                            string PartNo = worksheet.Cells[row, 1].Text.ToUpper().Trim();
                            string Model = worksheet.Cells[row, 2].Text.ToUpper().Trim();
                            string newConfirm = worksheet.Cells[row, 6].Text.Trim();

                            if (String.IsNullOrEmpty(PartNo)) break;

                            // lưu dữ liệu vào database
                            var checkExist = db.Dispose_ServicePart.Where(x => x.PartNo == PartNo).FirstOrDefault();
                            if (checkExist == null)
                            {
                                fail++;
                                listFail.Add("Not exist " + PartNo + " in service part list");
                                temMsg = "Not exist";
                            }
                            else
                            {
                                PDCConfirmServicePart(checkExist.ServicePartID, newConfirm, Model);
                                suc++;
                                temMsg = "Updated";
                            }

                            worksheet.Cells[row, 7].Value = temMsg;
                        }
                        catch
                        {

                        }

                    }
                    package.Save();
                    status = true;
                }
            }

            else
            {
                msg = "You do not permission!";
            }

            return Json(new { status = status, msg = "Success " + suc + " & fail " + fail + ". Detail check list upload control." }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult PDCConfirmServicePart(int ServicePartID, string newConfirm, string model)
        {
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            var name = Session["Name"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            bool status = false;
            string msg = "";
            var today = DateTime.Now;
            var checkExist = db.Dispose_ServicePart.Find(ServicePartID);
            if (dept == "PDC" && (role == "Check" || role == "Approve") || admin == "Admin")
            {
                // lưu dữ liệu vào database

                if (checkExist == null)
                {
                    return Json(new { status = false, msg = "No exist this item!" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    checkExist.Model = String.IsNullOrEmpty(model) ? checkExist.Model : model;
                    checkExist.LatestConfirm = newConfirm?.ToUpper();
                    checkExist.ConfirmBy = name;
                    checkExist.ConfirmDate = today;
                    db.Entry(checkExist).State = EntityState.Modified;
                    db.SaveChanges();
                    status = true;
                }
            }
            else
            {
                msg = "You do not have permission!";
            }
            return Json(new { status = status, msg = msg, data = checkExist }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult ExportExcelListServicePart()
        {

            string handle = Guid.NewGuid().ToString();
            MemoryStream output = new MemoryStream();
            var listService = db.Dispose_ServicePart.ToList();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/DisposeDie/Format/ConfirmServicePart.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();
                int rowId = 4;

                sheet.Cells["F1"].Value = "Date: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                foreach (var x in listService)
                {
                    sheet.Cells["A" + rowId.ToString()].Value = x.PartNo;
                    sheet.Cells["B" + rowId.ToString()].Value = x.Model;
                    sheet.Cells["C" + rowId.ToString()].Value = x.LatestConfirm;
                    sheet.Cells["D" + rowId.ToString()].Value = x.ConfirmBy;
                    sheet.Cells["E" + rowId.ToString()].Value = x.ConfirmDate;
                    rowId++;
                }
                package.SaveAs(output);
                output.Position = 0;
                TempData[handle] = output.ToArray();

            }
            var result = new { FileGuid = handle, FileName = DateTime.Now.ToString("yyyyMMdd-HHmmss") + "_List_Part_Service.xlsx" };

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getListEOLPart()
        {
            var output = db.Dispose_UniquePart.OrderByDescending(x => x.UniquePartID).ToList();
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getListServicePart()
        {
            var output = db.Dispose_ServicePart.OrderByDescending(x => x.ServicePartID).ToList();
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult uploadInventResult(HttpPostedFileBase file)
        {
            bool status = false;
            string msg = "";
            if (file == null)
            {
                status = false;
                msg = "Please select file inventory Result";
                return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
            }
            verifyForm verifyForm = isCorrectForm(file);
            if (verifyForm.isPass)
            {
                string mail = Session["Mail"].ToString();
                string name = Session["Name"].ToString();
                string dept = Session["Dept"].ToString();

                // save file vật lí
                string fileName = "InventoryResult-uploadDate-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
                string fileExt = Path.GetExtension(file.FileName);
                string path = Server.MapPath("~/File/DisposeDie/");
                fileName += fileExt;
                file.SaveAs(path + fileName);

                Thread t = new Thread(() =>
                {
                    commoneFunc.readFilesInventResult(path + fileName, mail, name, dept);
                });
                t.Start();
                t.IsBackground = true;
                status = true;
                msg = "System will take several time to process you file depend on your file size. So system will send you email after finish";
            }
            else
            {
                status = false;
                msg = verifyForm.msg;
            }

            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getFileNameUpload(string type)
        {
            db.Configuration.ProxyCreationEnabled = false;
            var output = db.Dispose_ControlFileUpload.Where(x => x.Type.Contains(type) && x.Active == true).OrderByDescending(x => x.UploadDate).Take(20).ToList();
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public class verifyForm
        {
            public bool isPass { set; get; }
            public string msg { set; get; }
        }
        public verifyForm isCorrectForm(HttpPostedFileBase file)
        {

            // 1. Check format đúng version chưa
            var A1 = "FORMART UPLOAD INVENTORY RESULT TO DMS";
            var A5 = "***";
            var C4 = "ASSET NO";
            var O4 = "Die ID (for die only)";
            var V4 = "Running Shot";
            var Y4 = "Is Match DMS DB";
            var msg = "";
            bool isPass = true;

            using (ExcelPackage package = new ExcelPackage(file.InputStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                // Check version
                var a1 = worksheet.Cells["A1"].Text.Trim();
                var a5 = worksheet.Cells["A5"].Text.Trim();
                var c4 = worksheet.Cells["C4"].Text.Trim();
                var o4 = worksheet.Cells["O4"].Text.Trim();
                var v4 = worksheet.Cells["V4"].Text.Trim();
                var y4 = worksheet.Cells["Y4"].Text.Trim();

                if (A1 != a1 || A5 != a5 || C4 != c4 || O4 != o4 || V4 != v4 || Y4 != y4)
                {
                    msg = "file: " + file.FileName + " is Wrong format!";
                    isPass = false;
                }

            }

            verifyForm output = new verifyForm()
            {
                isPass = isPass,
                msg = msg
            };
            return output;
        }


        public ActionResult ExportListInventoryResult(List<InventoryResult> invents)
        {
            MemoryStream output = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/DisposeDie/Format/FormInventoryResult.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();
                int rowId = 6;
                int i = 1;
                foreach (var item in invents)
                {
                    sheet.Cells["A" + rowId.ToString()].Value = item.InventoryID;
                    sheet.Cells["B" + rowId.ToString()].Value = i;
                    sheet.Cells["C" + rowId.ToString()].Value = item.AssetNo;
                    sheet.Cells["D" + rowId.ToString()].Value = item.OldAssetNo;
                    sheet.Cells["E" + rowId.ToString()].Value = item.CostCenter;
                    sheet.Cells["F" + rowId.ToString()].Value = item.DeptName;
                    sheet.Cells["G" + rowId.ToString()].Value = item.AssetName;
                    sheet.Cells["H" + rowId.ToString()].Value = item.Note1;
                    sheet.Cells["I" + rowId.ToString()].Value = item.ComponentAsset;
                    sheet.Cells["J" + rowId.ToString()].Value = item.ClassCode;
                    sheet.Cells["K" + rowId.ToString()].Value = item.OriginalCostUSD;
                    sheet.Cells["L" + rowId.ToString()].Value = item.RemainCostUSD;
                    sheet.Cells["M" + rowId.ToString()].Value = item.StartUseDate.HasValue ? String.Format("{0:MM/dd/yyyy}", item.StartUseDate) : "-";
                    sheet.Cells["N" + rowId.ToString()].Value = item.LocationSupplierCode;
                    sheet.Cells["O" + rowId.ToString()].Value = item.DieNo;
                    sheet.Cells["P" + rowId.ToString()].Value = item.IsFixAsset == true ? "Y" : "N";
                    sheet.Cells["Q" + rowId.ToString()].Value = item.FAPlate;
                    sheet.Cells["R" + rowId.ToString()].Value = item.UsingStatus;
                    sheet.Cells["S" + rowId.ToString()].Value = item.StopDate.HasValue ? String.Format("{0:MM/dd/yyyy}", item.StopDate) : "-";
                    sheet.Cells["T" + rowId.ToString()].Value = item.ActionPlanForUnuse;
                    sheet.Cells["U" + rowId.ToString()].Value = item.ReasonforDispose;
                    sheet.Cells["V" + rowId.ToString()].Value = item.Shot;
                    sheet.Cells["W" + rowId.ToString()].Value = item.RecordShotDate.HasValue ? String.Format("{0:MM/dd/yyyy}", item.RecordShotDate) : "-";
                    sheet.Cells["X" + rowId.ToString()].Value = item.IsWrongLocation == true ? "Y" : "N";
                    sheet.Cells["Y" + rowId.ToString()].Value = item.IsMatchDMSDatabase == true ? "Y" : "N";
                    sheet.Cells["Z" + rowId.ToString()].Value = item.isVerifiedDisposeLastQuater == true ? "Y" : "N";
                    sheet.Cells["AA" + rowId.ToString()].Value = item.isNoneOperationFollowInventory == true ? "Y" : "N";
                    sheet.Cells["AB" + rowId.ToString()].Value = item.isNoneOperationFollowDMSCheck == true ? "Y" : "N";
                    sheet.Cells["AC" + rowId.ToString()].Value = item.IsSeletedForVerify == true ? "Y" : (item.isReferDie == true ? "REFER" : "N");
                    sheet.Cells["AD" + rowId.ToString()].Value = item.WarningContent;
                    i++;
                    rowId++;
                }

                package.SaveAs(output);
            }
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=InventoryResult_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx");
            output.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();
            return RedirectToAction("Index");
        }

        public JsonResult getSummarizeVerify()
        {
            var output = storeProcudure.GetSummarizeVerifyDieDispose();
            return Json(output, JsonRequestBehavior.AllowGet);
        }


        public JsonResult getSumarizeAllItemIsVerifying()
        {
            var output = storeProcudure.getSumarizeAllItemIsVerifying();
            var quaterOPEN = db.Dispose_Quater.Where(x => x.Status == "OPEN").Select(x => new
            {
                Quater = x.Quater,
                QuaterID = x.QuaterID,
                Status = x.Status
            }).ToList();
            return Json(new { quaterOPEN = quaterOPEN, sum = output }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult MakeAnouncement(string[] deptReciept, string msg, string infor, int? quaterID, string quaterName)
        {
            var name = Session["Name"].ToString();
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            var status = false;
            if (admin == "Admin" || (dept == "CAM" && role == "Check"))
            {
                var today = DateTime.Now;
                int QID = quaterID != null ? (int)quaterID : 0;
                // ******* Create verified time
                if (!String.IsNullOrWhiteSpace(quaterName))
                {
                    QID = CreateQuater(quaterName);
                }
                if (QID == 0) goto Exit;
                // ******* Chuyen sang verify page 
                var listSelected = db.InventoryResults.Where(x => (x.IsSeletedForVerify == true || x.isReferDie == true) && x.Active != false).ToList();
                foreach (var item in listSelected)
                {
                    // Kiểm tra item này có ĐANG trong gói thẩm định hay ko?
                    var checkExist = db.VerifyDisposeDies.Where(x => x.AssetNo == item.AssetNo && x.DieNo == item.DieNo && x.Dispose_Quater.Status == "OPEN").FirstOrDefault();
                    // Nếu chưa có thì add vào list thẩm định.
                    // Nếu có rồi thì ko add nữa.
                    if (checkExist == null)
                    {
                        VerifyDisposeDie newItem = new VerifyDisposeDie();
                        {
                            newItem.DISNo = genarateDISNo(QID);
                            newItem.QuaterID = QID;
                            newItem.DieNo = item.DieNo?.Trim();
                            newItem.AssetNo = item.AssetNo?.Trim();
                            newItem.InventoryID = item.InventoryID;
                            newItem.DieBelong = db.Die1.Where(x => x.DieNo == item.DieNo && x.Active != false).FirstOrDefault()?.Belong;
                            newItem.CreateBy = name;
                            newItem.CreateDate = today;
                            // add 22.May.2024
                            newItem.isReferDie = item.isReferDie;
                        }
                        db.VerifyDisposeDies.Add(newItem);
                        db.SaveChanges();
                    }
                }
                status = true;

                // ****** Gui mail to Related Dept
                sendEmailJob.sendEmainAnnounceVerifyDisposeDie(deptReciept, db.Dispose_Quater.Find(QID).Quater, msg, infor, Session["Mail"].ToString());
            }

        Exit:
            return (Json(status, JsonRequestBehavior.AllowGet));
        }

        public string genarateDISNo(int QID)
        {
            string DISNo = "";
            var quater = db.Dispose_Quater.Find(QID).Quater;
            int count = db.VerifyDisposeDies.Where(x => x.QuaterID == QID).Count() + 1;
            DISNo = quater + "-" + count.ToString();
            return DISNo;
        }

        public JsonResult addItemVerify(int currentVerifyID, int DieIDAdded)
        {


            bool status = false;
            string msg = "";
            var die = db.Die1.Find(DieIDAdded);
            var currentVerifyItem = db.VerifyDisposeDies.Find(currentVerifyID);

            bool isOpen = db.Dispose_Quater.Find(currentVerifyItem.QuaterID).Status == "OPEN";
            if (isOpen)
            {
                // Lấy current fileInventory ID 
                var fileInventID = db.InventoryResults.Find(currentVerifyItem.InventoryID)?.FileInventID;

                var existInventory = db.InventoryResults.Where(x => x.DieNo == die.DieNo && x.FileInventID == fileInventID).FirstOrDefault();
                if (existInventory == null)
                {
                    existInventory = commoneFunc.autoAddInventoryResult(die, (int)fileInventID, Session["Name"].ToString());
                }
                // Kiểm tra item này có ĐANG trong gói thẩm định hay ko?
                var checkExist = db.VerifyDisposeDies.Where(x => x.DieNo == die.DieNo && x.Dispose_Quater.Status == "OPEN").FirstOrDefault();
                // Nếu chưa có thì add vào list thẩm định.

                if (checkExist == null)
                {
                    VerifyDisposeDie newItem = new VerifyDisposeDie();
                    {
                        newItem.DISNo = genarateDISNo(currentVerifyItem.QuaterID);
                        newItem.QuaterID = currentVerifyItem.QuaterID;
                        newItem.DieNo = die.DieNo?.Trim();
                        newItem.AssetNo = die.FixedAssetNo?.Trim();
                        newItem.InventoryID = (int)(existInventory?.InventoryID);
                        newItem.DieBelong = die?.Belong;
                        newItem.CreateBy = Session["Name"].ToString();
                        newItem.CreateDate = DateTime.Now;
                        newItem.isReferDie = true;
                    }
                    db.VerifyDisposeDies.Add(newItem);
                    db.SaveChanges();
                    status = true;
                }
                else   // Nếu có rồi thì ko add nữa.
                {
                    status = false;
                    msg = "Already exist this item!";
                }
            }
            else
            {
                status = false;
                msg = "Bạn ko thêm được die này vào verify do bạn đang chọn gói thẩm định đã CLOSE.";
            }

            return Json(new { status = status, msg = msg }, JsonRequestBehavior.AllowGet);
        }

        public int CreateQuater(string quaterName)
        {
            int id = 0;
            var dept = Session["Dept"].ToString();
            // string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            if (dept.ToUpper() == "CAM" || admin == "Admin")
            {
                if (!String.IsNullOrWhiteSpace(quaterName))
                {
                    var ExistQ = db.Dispose_Quater.Where(x => x.Quater.Contains(quaterName)).FirstOrDefault();
                    if (ExistQ == null)
                    {
                        Dispose_Quater newQuater = new Dispose_Quater()
                        {
                            Quater = quaterName.ToUpper(),
                            Status = "OPEN"
                        };
                        db.Dispose_Quater.Add(newQuater);
                        db.SaveChanges();
                        id = newQuater.QuaterID;
                    }
                    else
                    {
                        ExistQ.Status = "OPEN";
                        db.Entry(ExistQ).State = EntityState.Modified;
                        db.SaveChanges();
                        id = ExistQ.QuaterID;
                    }


                }
            }
            return id;
        }

        public JsonResult CheckProgress(int quaterID)
        {
            var allItems = db.VerifyDisposeDies.Where(x => x.QuaterID == quaterID).ToList();
            int DoneInput = allItems.Where(x => !String.IsNullOrEmpty(x.DieBelong)
               && x.PUR_TotalCav != null && x.PUR_CycleTime != null
                && !String.IsNullOrEmpty(x.PAE_CRG_DMT_PUC_Remark)
                && x.PUC_WarrantyShot != null
                && !String.IsNullOrEmpty(x.PE_AllModel)
                && !String.IsNullOrEmpty(x.PE_CommonPart_NewModel)
                && !String.IsNullOrEmpty(x.PE_CommonPart_CurrentModel)
                && !String.IsNullOrEmpty(x.PE_AlternativePart)
                && !String.IsNullOrEmpty(x.PE_FamilyPart)
                && !String.IsNullOrEmpty(x.PDC_PartDemandStatus)
                && x.PDC_MP_MaxDemand != null
                && x.PDC_MP_MaxDemand != null
                ).Count();

            int DoneVerify = allItems.Where(x => !String.IsNullOrWhiteSpace(x.CAM_KeepOrDispose)).Count();
            int NoVerifyDispose = allItems.Where(x => x.CAM_KeepOrDispose == "DISPOSE").Count();
            int NoVerifyKeep = allItems.Where(x => x.CAM_KeepOrDispose == "KEEP").Count();
            int DoneServey = allItems.Where(x => !String.IsNullOrWhiteSpace(x.PUC_DMT_ConfirmKeepOrDispose)).Count();
            int ServeyDispose = allItems.Where(x => x.PUC_DMT_ConfirmKeepOrDispose == "DISPOSE").Count();
            int ServeyKeep = allItems.Where(x => x.PUC_DMT_ConfirmKeepOrDispose == "KEEP").Count();
            int DoneDecision = allItems.Where(x => x.TopApproveDate != null).Count();
            int DecisionDispose = allItems.Where(x => x.TopApproveDate != null && x.FinalDecisionID == 2).Count();
            int DecisionKeep = allItems.Where(x => x.TopApproveDate != null && x.FinalDecisionID == 1).Count();
            int DonePhy = allItems.Where(x => x.FinalDecisionID == 2 && x.PhysicalDisposeDate != null).Count();
            int total = allItems.Count();

            List<object> output = new List<object>();
            //1 select
            output.Add(
                new
                {
                    ProcessName = "Selected",
                    Total = total,
                    Select = allItems.Where(x => x.isReferDie != true).Count(),
                    Refer = allItems.Where(x => x.isReferDie == true).Count()
                }
                );
            //2 Input
            output.Add(
                new
                {
                    ProcessName = "Input Information",
                    Total = total,
                    Done = DoneInput
                }
                );

            //3 CAM Verify
            output.Add(
                new
                {
                    ProcessName = "Verify",
                    Total = total,
                    Done = DoneVerify,
                    Dispose = NoVerifyDispose,
                    Keep = NoVerifyKeep
                }
                );
            //4 Servey
            output.Add(
                new
                {
                    ProcessName = "Servey",
                    Total = total,
                    Done = DoneServey,
                    Dispose = ServeyDispose,
                    Keep = ServeyKeep
                }
                );

            //5 Decision
            output.Add(
                new
                {
                    ProcessName = "Decision",
                    Total = total,
                    Done = DoneDecision,
                    Dispose = DecisionDispose,
                    Keep = DecisionKeep
                }
                );

            //5 Decision
            output.Add(
                new
                {
                    ProcessName = "Physical",
                    Total = DecisionDispose,
                    Done = DonePhy,
                }
                );

            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getLatestVerified(string assetNo, string dieNo)
        {
            var output = storeProcudure.getLatestDisposalVerified(assetNo, dieNo);
            return Json(output, JsonRequestBehavior.AllowGet);
        }

        public JsonResult getSumarizeDeptResponse()
        {
            var output = storeProcudure.GetSummarizeDeptResponseVerifyDispose();

            return Json(output, JsonRequestBehavior.AllowGet);
        }


        public class waringContent
        {
            public string dept { set; get; }
            public string id { set; get; }
            public string waring { set; get; }
        }

        public class dataexport
        {
            public string VerifyTime { set; get; }
            public string DISNo { set; get; }
            public waringContent[] DeptAction { set; get; }
            public string DieRefer { set; get; }
            public string GenaralInfor { set; get; }
            public string AssetType { set; get; }
            public string AssetNo { set; get; }
            public string OldAssetNo { set; get; }
            public string FACostCenter { set; get; }
            public string SupplierCode { set; get; }
            public string VietName_OverSea { set; get; }
            public string OriginalCost { set; get; }
            public string RemainCost { set; get; }
            public string AssetName { set; get; }
            public string PartNo { set; get; }
            public string DieNo { set; get; }
            public string DieDim { set; get; }
            public string DieCategory { set; get; }
            public string ProcessType { set; get; }
            public string TotalDie { set; get; }
            public string ToTalInUse { set; get; }
            public string StartUse { set; get; }
            public string UsingStatus { set; get; }
            public string StopDate { set; get; }
            public string UsingTime { set; get; }
            public string StopTime { set; get; }
            public string RequestDispose { set; get; }
            public string ReasonForDispose { set; get; }
            public string TotalCav { set; get; }
            public string CycleTime { set; get; }
            public string MonthCapa { set; get; }
            public string DieStatus { set; get; }
            public string RepairHistory { set; get; }
            public string RepairCode { set; get; }
            public string PAE_CRG_DMT_PUC_Comment { set; get; }
            public string GuaranteeShot { set; get; }
            public string Shot { set; get; }
            public string AllModel { set; get; }
            public string CommonPartOld { set; get; }
            public string CommonPartNew { set; get; }
            public string AlternativePart { set; get; }
            public string FamilyPart { set; get; }
            public string PartDemandStatus { set; get; }
            public string MPMaxDemand { set; get; }
            public string JPFeedback { set; get; }
            public string PDCRemark { set; get; }
            public string DieDemand { set; get; }
            public string Last_VerifyResult { set; get; }
            public string Last_VerifyReason { set; get; }
            public string Last_Quater { set; get; }
            public string DMS_Result { set; get; }
            public string DMS_Concept { set; get; }
            public string CAM_KeepOrDispose { set; get; }
            public string CAM_ReasonDecision { set; get; }
            public string CAM_NextStep { set; get; }
            public string CAM_DeptInCharge { set; get; }
            public string CAM_Remark { set; get; }
            public string S_KeepOrDispose { set; get; }
            public string S_KeepOrDisposeReason { set; get; }
            public string S_Common_Family { set; get; }
            public string S_DMSWarning { set; get; }
            public string S_ReasonNotMatch { set; get; }
            public string CAM_PreDecision { set; get; }
            public string CAM_FinalDecision { set; get; }
            public string CAM_TOPApproveDate { set; get; }
            public string PhysicalDisposeDate { set; get; }
            public string ReasonChangeDecision { set; get; }
            public bool isReferDie { set; get; }
        }

        public string[] renderWaring(waringContent[] x)
        {
            string content = "";
            string dept = "";
            if (x != null)
            {
                foreach (var item in x)
                {
                    content += item.dept + ": " + item.waring + System.Environment.NewLine;
                    dept += item.dept + ",";
                }
            }

            string[] output = { dept, content };
            return output;
        }
        public JsonResult exportToExcel(List<dataexport> data)
        {

            string handle = Guid.NewGuid().ToString();
            MemoryStream output = new MemoryStream();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(Server.MapPath("~/File/DisposeDie/Format/FormVerificationDetail.xlsx"))))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.First();
                int rowId = 8;
                int i = 1;
                sheet.Cells["A3"].Value = "Date: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                if (data != null)
                {
                    foreach (var x in data)
                    {
                        sheet.Cells["A" + rowId.ToString()].Value = i;
                        sheet.Cells["B" + rowId.ToString()].Value = x.DISNo;
                        sheet.Cells["C" + rowId.ToString()].Value = renderWaring(x.DeptAction)[0];
                        sheet.Cells["D" + rowId.ToString()].Value = x.GenaralInfor;

                        sheet.Cells["E" + rowId.ToString()].Value = x.AssetType;
                        sheet.Cells["F" + rowId.ToString()].Value = x.AssetNo;
                        sheet.Cells["G" + rowId.ToString()].Value = x.OldAssetNo;
                        sheet.Cells["H" + rowId.ToString()].Value = x.FACostCenter;
                        sheet.Cells["I" + rowId.ToString()].Value = x.SupplierCode;
                        sheet.Cells["J" + rowId.ToString()].Value = x.VietName_OverSea;
                        sheet.Cells["K" + rowId.ToString()].Value = x.OriginalCost;
                        sheet.Cells["L" + rowId.ToString()].Value = x.RemainCost;
                        sheet.Cells["M" + rowId.ToString()].Value = x.AssetName;
                        sheet.Cells["N" + rowId.ToString()].Value = x.PartNo;
                        sheet.Cells["O" + rowId.ToString()].Value = x.DieNo;
                        sheet.Cells["P" + rowId.ToString()].Value = x.DieDim;
                        sheet.Cells["Q" + rowId.ToString()].Value = x.DieCategory;
                        sheet.Cells["R" + rowId.ToString()].Value = x.ProcessType;
                        sheet.Cells["S" + rowId.ToString()].Value = x.ToTalInUse + "/" + x.TotalDie;
                        sheet.Cells["T" + rowId.ToString()].Value = x.StartUse;
                        sheet.Cells["U" + rowId.ToString()].Value = x.UsingStatus;
                        sheet.Cells["V" + rowId.ToString()].Value = x.StopDate;
                        sheet.Cells["W" + rowId.ToString()].Value = x.UsingTime;
                        sheet.Cells["X" + rowId.ToString()].Value = x.StopTime;
                        sheet.Cells["Y" + rowId.ToString()].Value = x.RequestDispose;
                        sheet.Cells["Z" + rowId.ToString()].Value = x.ReasonForDispose;
                        sheet.Cells["AA" + rowId.ToString()].Value = x.TotalCav;
                        sheet.Cells["AB" + rowId.ToString()].Value = x.CycleTime;
                        sheet.Cells["AC" + rowId.ToString()].Value = x.MonthCapa;
                        sheet.Cells["AD" + rowId.ToString()].Value = x.DieStatus;
                        sheet.Cells["AE" + rowId.ToString()].Value = x.RepairHistory;
                        sheet.Cells["AF" + rowId.ToString()].Value = x.PAE_CRG_DMT_PUC_Comment;
                        sheet.Cells["AG" + rowId.ToString()].Value = x.GuaranteeShot;
                        sheet.Cells["AH" + rowId.ToString()].Value = x.Shot;
                        sheet.Cells["AI" + rowId.ToString()].Value = x.AllModel;
                        sheet.Cells["AJ" + rowId.ToString()].Value = x.CommonPartOld;
                        sheet.Cells["AK" + rowId.ToString()].Value = x.CommonPartNew;
                        sheet.Cells["AL" + rowId.ToString()].Value = x.AlternativePart;
                        sheet.Cells["AM" + rowId.ToString()].Value = x.FamilyPart;
                        sheet.Cells["AN" + rowId.ToString()].Value = x.PartDemandStatus;
                        sheet.Cells["AO" + rowId.ToString()].Value = x.MPMaxDemand;
                        sheet.Cells["AP" + rowId.ToString()].Value = x.JPFeedback;
                        sheet.Cells["AQ" + rowId.ToString()].Value = x.PDCRemark;
                        sheet.Cells["AR" + rowId.ToString()].Value = x.DieDemand;
                        sheet.Cells["AS" + rowId.ToString()].Value = x.Last_VerifyResult;
                        sheet.Cells["AT" + rowId.ToString()].Value = x.Last_VerifyReason;
                        sheet.Cells["AU" + rowId.ToString()].Value = x.Last_Quater;
                        sheet.Cells["AV" + rowId.ToString()].Value = x.DMS_Result;
                        sheet.Cells["AW" + rowId.ToString()].Value = x.DMS_Concept;
                        sheet.Cells["AX" + rowId.ToString()].Value = x.CAM_KeepOrDispose;
                        sheet.Cells["AY" + rowId.ToString()].Value = x.CAM_ReasonDecision;
                        sheet.Cells["AZ" + rowId.ToString()].Value = x.CAM_NextStep;
                        sheet.Cells["BA" + rowId.ToString()].Value = x.CAM_DeptInCharge;
                        sheet.Cells["BB" + rowId.ToString()].Value = x.CAM_Remark;
                        sheet.Cells["BC" + rowId.ToString()].Value = x.S_KeepOrDispose;
                        sheet.Cells["BD" + rowId.ToString()].Value = x.S_KeepOrDisposeReason;
                        sheet.Cells["BE" + rowId.ToString()].Value = x.S_Common_Family;
                        sheet.Cells["BF" + rowId.ToString()].Value = x.S_DMSWarning;
                        sheet.Cells["BG" + rowId.ToString()].Value = x.S_ReasonNotMatch;
                        sheet.Cells["BH" + rowId.ToString()].Value = x.CAM_PreDecision;
                        sheet.Cells["BI" + rowId.ToString()].Value = x.CAM_FinalDecision;
                        sheet.Cells["BJ" + rowId.ToString()].Value = x.CAM_TOPApproveDate;
                        sheet.Cells["BK" + rowId.ToString()].Value = x.PhysicalDisposeDate;
                        sheet.Cells["BL" + rowId.ToString()].Value = x.ReasonChangeDecision;
                        sheet.Cells["BM" + rowId.ToString()].Value = x.isReferDie == true ? "Y" : "N";
                        i++;
                        rowId++;
                    }
                }

                package.SaveAs(output);
                package.Workbook.Calculate();
                output.Position = 0;
                TempData[handle] = output.ToArray();
            }

            var result = new { FileGuid = handle, FileName = DateTime.Now.ToString("yyyyMMdd-HHmmss") + "_Die_Dispose_Verify.xlsx" };

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ImportFromExcel(HttpPostedFileBase file)
        {
            var name = Session["Name"].ToString();
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            bool status = false;
            List<string> success = new List<string>();
            List<string> fail = new List<string>();
            string handle = Guid.NewGuid().ToString();
            if (role == "Check" || admin == "Admin")
            {
                // Doc file

                MemoryStream output = new MemoryStream();
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    // Check form
                    var B6 = worksheet.Cells["B6"].Text; // Control No
                    var F6 = worksheet.Cells["F6"].Text; // Asset No
                    var BJ6 = worksheet.Cells["BJ6"].Text; // TOP approve Date


                    if (B6 != "ControlNo" || F6 != "Asset No" || BJ6 != "TOP approve Date")
                    {
                        fail.Add("Sai Format, Format đã bị thêm/Xóa cột");
                        var data1 = new
                        {
                            success = success.Count(),
                            fail = fail
                        };
                        return Json(data1, JsonRequestBehavior.AllowGet);
                    }
                    // Kết thúc check form

                    for (int row = start.Row + 7; row <= end.Row; row++)
                    { // Row by row...
                        var DISNo = worksheet.Cells[row, 2].Text;
                        var msg = "";
                        if (String.IsNullOrEmpty(DISNo)) break;
                        var assetNo = worksheet.Cells[row, 6].Text.Trim();
                        var findItem = db.VerifyDisposeDies.Where(x => x.DISNo == DISNo && x.AssetNo == assetNo).FirstOrDefault();

                        if (findItem == null)
                        {
                            msg = "Không tìm thấy item có Control No : " + DISNo + " cho Asset " + assetNo;
                            goto exitLoop;
                        }

                        int id = findItem.VerifyID;
                        string value = "";



                        //All Dept
                        {
                            value = worksheet.Cells[row, 4].Text;
                            if (!(findItem.GenaralInfor == value || (String.IsNullOrEmpty(findItem.GenaralInfor) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.GenaralInfor) == true)))
                            {
                                if (value.Contains(findItem.GenaralInfor))
                                {
                                    value = value.Replace(findItem.GenaralInfor, "");
                                    var test = saveData(id, "GenaralInfor", value);
                                    msg += "[#4]Genaral Information: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                                else
                                {
                                    msg += "[#4]Genaral Information: " + "{status = false, msg = Please do not delete others comment}" + System.Environment.NewLine;
                                }
                            }


                        }

                        if (findItem.DieBelong != "LBP" && findItem.DieBelong != "CRG")
                        {
                            {
                                value = worksheet.Cells[row, 30].Text;
                                if (!(findItem.Manual_DieStatus == value || (String.IsNullOrEmpty(findItem.Manual_DieStatus) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.Manual_DieStatus) == true)))
                                {
                                    var test = saveData(id, "Manual_DieStatus", value);
                                    msg += "[#30]Die Status: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 31].Text;
                                if (!(findItem.Manual_DieRepairHistory == value || (String.IsNullOrEmpty(findItem.Manual_DieRepairHistory) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.Manual_DieRepairHistory) == true)))
                                {
                                    var test = saveData(id, "Manual_DieRepairHistory", value);
                                    msg += "[#31]History of Repair/Modify: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 19].Text;
                                int qtyDie = 0;
                                bool isNum = int.TryParse(value, out qtyDie);
                                if (!(findItem.Manual_QtyOfDie == qtyDie || (String.IsNullOrEmpty(findItem.Manual_QtyOfDie?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "Manual_QtyOfDie", value);
                                    msg += "[#19]Qty of Die: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 29].Text;
                                double cap = 0;
                                bool isNum = double.TryParse(value, out cap);
                                if (!(findItem.Manual_Capacity == cap || (String.IsNullOrEmpty(findItem.Manual_Capacity?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "Manual_QtyOfDie", value);
                                    msg += "[#29]Month Capacity : " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }
                        }



                        // CAM

                        if (dept == "CAM" || admin == "Admin")
                        {
                            {
                                value = worksheet.Cells[row, 17].Text;
                                if (!(findItem.DieBelong == value || (String.IsNullOrEmpty(findItem.DieBelong) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.DieBelong) == true)))
                                {
                                    var test = saveData(id, "DieBelong", value);
                                    msg += "[#17]Die Category: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 50].Text;
                                if (!(findItem.CAM_KeepOrDispose == value || (String.IsNullOrEmpty(findItem.CAM_KeepOrDispose) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.CAM_KeepOrDispose) == true)))
                                {
                                    var test = saveData(id, "CAM_KeepOrDispose", value);
                                    msg += "[#50]Keep or Dispose: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 51].Text;
                                if (!(findItem.CAM_ReasonOfDecision == value || (String.IsNullOrEmpty(findItem.CAM_ReasonOfDecision) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.CAM_ReasonOfDecision) == true)))
                                {
                                    var test = saveData(id, "CAM_ReasonOfDecision", value);
                                    msg += "[#51]Reason of decision: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 52].Text;
                                if (!(findItem.CAM_NextVerifyStep == value || (String.IsNullOrEmpty(findItem.CAM_NextVerifyStep) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.CAM_NextVerifyStep) == true)))
                                {
                                    var test = saveData(id, "CAM_NextVerifyStep", value);
                                    msg += "[#52]Next Verify Step: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 53].Text;
                                if (!(findItem.CAM_DeptInCharge == value || (String.IsNullOrEmpty(findItem.CAM_DeptInCharge) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.CAM_DeptInCharge) == true)))
                                {
                                    var test = saveData(id, "CAM_DeptInCharge", value);
                                    msg += "[53]Dept In Charge: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 54].Text;
                                if (!(findItem.CAM_Remak == value || (String.IsNullOrEmpty(findItem.CAM_Remak) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.CAM_Remak) == true)))
                                {
                                    var test = saveData(id, "CAM_Remak", value);
                                    msg += "[#54]CAM Remak: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 60].Text;

                                if (!(db.Dispose_DecisionCategories.Find(findItem.PreDecisionID)?.Type == value || (String.IsNullOrEmpty(findItem.PreDecisionID?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "Pre_Decision", value);
                                    msg += "[#60]Pre-Decision: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 61].Text;

                                if (!(db.Dispose_DecisionCategories.Find(findItem.FinalDecisionID)?.Type == value || (String.IsNullOrEmpty(findItem.FinalDecisionID?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "Final_Decision", value);
                                    msg += "[#61]Final Decision: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }
                            {
                                value = worksheet.Cells[row, 62].Text;
                                DateTime datetime = new DateTime(2000, 1, 1);
                                bool r = DateTime.TryParse(value, out datetime);
                                //var s = 
                                if (!(findItem.TopApproveDate?.ToString("yyyyMMdd") == datetime.ToString("yyyyMMdd") || (String.IsNullOrEmpty(findItem.TopApproveDate?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "TopApproveDate", value);
                                    msg += "[#62]Top Approve Date: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 63].Text;
                                DateTime datetime = new DateTime(2000, 1, 1);
                                bool r = DateTime.TryParse(value, out datetime);
                                //var s = 
                                if (!(findItem.PhysicalDisposeDate?.ToString("yyyyMMdd") == datetime.ToString("yyyyMMdd") || (String.IsNullOrEmpty(findItem.PhysicalDisposeDate?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "PhysicalDisposeDate", value);
                                    msg += "[#63]PhysicalDisposeDate: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }




                        }

                        if (dept == "PUR" || admin == "Admin")
                        {
                            {
                                value = worksheet.Cells[row, 27].Text;
                                int cav = 0;
                                bool isNum = int.TryParse(value, out cav);
                                if (!(findItem.PUR_TotalCav == cav || (String.IsNullOrEmpty(findItem.PUR_TotalCav?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "PUR_TotalCav", value);
                                    msg += "[#27]Total Cavity: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 28].Text;
                                double ct = 0;
                                bool isNum = double.TryParse(value, out ct);
                                if (!(findItem.PUR_CycleTime == ct || (String.IsNullOrEmpty(findItem.PUR_CycleTime?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "PUR_CycleTime", value);
                                    msg += "[#28]Cycle Time (s): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                        }


                        if (dept == "PAE" || dept == "CRG" || dept == "DMT" || dept == "PUC" || admin == "Admin")
                        {
                            {
                                value = worksheet.Cells[row, 32].Text;
                                if (!(findItem.PAE_CRG_DMT_PUC_Remark == value || (String.IsNullOrEmpty(findItem.PAE_CRG_DMT_PUC_Remark) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PAE_CRG_DMT_PUC_Remark) == true)))
                                {
                                    var test = saveData(id, "PAE_CRG_DMT_PUC_Remark", value);
                                    msg += "[#32]PAE/CRG/DMT/PUC Re-check and Comment Die Status: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }
                        }

                        if (dept == "DMT" || dept == "PUC" || admin == "Admin")
                        {
                            {
                                value = worksheet.Cells[row, 33].Text;
                                double gr = 0;
                                bool isNum = double.TryParse(value, out gr);
                                if (!(findItem.PUC_WarrantyShot == gr || (String.IsNullOrEmpty(findItem.PUC_WarrantyShot?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "PUC_WarrantyShot", value);
                                    msg += "[#33]Guarantee shot: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 55].Text;
                                if (!(findItem.PUC_DMT_ConfirmKeepOrDispose == value || (String.IsNullOrEmpty(findItem.PUC_DMT_ConfirmKeepOrDispose) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PUC_DMT_ConfirmKeepOrDispose) == true)))
                                {
                                    var test = saveData(id, "PUC_DMT_ConfirmKeepOrDispose", value);
                                    msg += "[#55]Supplier/DMT confirm keep or dispose: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 56].Text;
                                if (!(findItem.PUC_DMT_ReasonKeepOrDispose == value || (String.IsNullOrEmpty(findItem.PUC_DMT_ReasonKeepOrDispose) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PUC_DMT_ReasonKeepOrDispose) == true)))
                                {
                                    var test = saveData(id, "PUC_DMT_ReasonKeepOrDispose", value);
                                    msg += "[#56]Reason keep or Dispose: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 57].Text;
                                if (!(findItem.PUC_DMT_ConfirmHasCommonOrFamily == value || (String.IsNullOrEmpty(findItem.PUC_DMT_ConfirmHasCommonOrFamily) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PUC_DMT_ConfirmHasCommonOrFamily) == true)))
                                {
                                    var test = saveData(id, "PUC_DMT_ConfirmHasCommonOrFamily", value);
                                    msg += "[#57]Using dept/Supplier confirm: have common part or not?: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 59].Text;
                                if (!(findItem.PUC_DMT_Reason_NotMatch == value || (String.IsNullOrEmpty(findItem.PUC_DMT_Reason_NotMatch) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PUC_DMT_Reason_NotMatch) == true)))
                                {
                                    var test = saveData(id, "PUC_DMT_Reason_NotMatch", value);
                                    msg += "[#59]Reason (not match): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }
                        }

                        if (dept.Contains("PE") || admin == "Admin")
                        {
                            if (findItem.isReferDie == true) goto exitLoop;
                            {
                                value = worksheet.Cells[row, 35].Text;
                                if (!(findItem.PE_AllModel == value || (String.IsNullOrEmpty(findItem.PE_AllModel) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PE_AllModel) == true)))
                                {
                                    var test = saveData(id, "PE_AllModel", value);
                                    msg += "[#35]All Model: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 36].Text;
                                if (!(findItem.PE_CommonPart_CurrentModel == value || (String.IsNullOrEmpty(findItem.PE_CommonPart_CurrentModel) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PE_CommonPart_CurrentModel) == true)))
                                {
                                    var test = saveData(id, "PE_CommonPart_CurrentModel", value);
                                    msg += "[#36]Common part with (A) (old & current model): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 37].Text;
                                if (!(findItem.PE_CommonPart_NewModel == value || (String.IsNullOrEmpty(findItem.PE_CommonPart_NewModel) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PE_CommonPart_NewModel) == true)))
                                {
                                    var test = saveData(id, "PE_CommonPart_NewModel", value);
                                    msg += "[#37]Common part with (A) (New model): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 38].Text;
                                if (!(findItem.PE_AlternativePart == value || (String.IsNullOrEmpty(findItem.PE_AlternativePart) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PE_AlternativePart) == true)))
                                {
                                    var test = saveData(id, "PE_AlternativePart", value);
                                    msg += "[#38]Alternative part with (A) (old, current & new models): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 39].Text;
                                if (!(findItem.PE_FamilyPart == value || (String.IsNullOrEmpty(findItem.PE_FamilyPart) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PE_FamilyPart) == true)))
                                {
                                    var test = saveData(id, "PE_FamilyPart", value);
                                    msg += "[#39]Family part with (A) (old, current & new models): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                        }

                        if (dept == "PDC" || admin == "Admin")
                        {
                            if (findItem.isReferDie == true) goto exitLoop;

                            {
                                value = worksheet.Cells[row, 40].Text;
                                if (!(findItem.PDC_PartDemandStatus == value || (String.IsNullOrEmpty(findItem.PDC_PartDemandStatus) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PDC_PartDemandStatus) == true)))
                                {
                                    var test = saveData(id, "PDC_PartDemandStatus", value);
                                    msg += "[#40]Part demand status: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 41].Text;
                                double dm = 0;
                                bool isNum = double.TryParse(value, out dm);
                                if (!(findItem.PDC_MP_MaxDemand == dm || (String.IsNullOrEmpty(findItem.PDC_MP_MaxDemand?.ToString()) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(value) == true)))
                                {
                                    var test = saveData(id, "PDC_MP_MaxDemand", value);
                                    msg += "[#41]MP Max demand/month (pcs): " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 42].Text;
                                if (!(findItem.PDC_ResultJP_FB == value || (String.IsNullOrEmpty(findItem.PDC_ResultJP_FB) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PDC_ResultJP_FB) == true)))

                                {
                                    var test = saveData(id, "PDC_ResultJP_FB", value);
                                    msg += "[#42]Result feedback from JP: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }

                            {
                                value = worksheet.Cells[row, 43].Text;
                                if (!(findItem.PDC_Remark == value || (String.IsNullOrEmpty(findItem.PDC_Remark) == String.IsNullOrEmpty(value) && String.IsNullOrEmpty(findItem.PDC_Remark) == true)))
                                {
                                    var test = saveData(id, "PDC_Remark", value);
                                    msg += "[#43] Remark: " + test.Data.ToString() + System.Environment.NewLine;
                                }
                            }
                        }


                    exitLoop:
                        worksheet.Cells[row, 65].Value = msg;
                    }
                    package.SaveAs(output);



                    // 1. Lưu file & control
                    // save file vật lí
                    string fileName = "Dept-Upload-File-Verify-Detail-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/DisposeDie/");
                    fileName += fileExt;
                    package.SaveAs(path + fileName);


                    Dispose_ControlFileUpload newFile = new Dispose_ControlFileUpload()
                    {
                        FileName = fileName,
                        Type = "Dept_Verify_Detail",
                        Dept = dept,
                        UploadBy = name,
                        UploadDate = DateTime.Now,
                        Active = true
                    };
                    db.Dispose_ControlFileUpload.Add(newFile);
                    db.SaveChanges();

                    output.Position = 0;
                    TempData[handle] = output.ToArray();
                    status = true;
                }
            }
            var data = new { status = status, FileGuid = handle, FileName = "Result_Upload_Disposal_Detail.xlsx" };
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ImportListCAMselectForVerify(HttpPostedFileBase file)
        {
            var name = Session["Name"].ToString();
            var dept = Session["Dept"].ToString();
            var role = Session["Dispose_Role"].ToString();
            string admin = Session["Code"].ToString() == "DISMEET" ? "Admin" : "";
            bool status = false;
            List<string> success = new List<string>();
            string msg = "";
            int selected = 0;
            string handle = Guid.NewGuid().ToString();
            if ((dept == "CAM" && role == "Check") || admin == "Admin")
            {
                MemoryStream output = new MemoryStream();
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    // Check form
                    var AC4 = worksheet.Cells["AC4"].Text; // Is Selected To Verify This Time
                    var A4 = worksheet.Cells["A4"].Text; // ID


                    if (AC4.Contains("Is Selected To Verify This Time") || !A4.Contains("Do not input or delete"))
                    {
                        msg = "Sai Format, Format đã bị thêm/Xóa cột";
                        var data1 = new
                        {
                            status = status,
                            success = success.Count(),
                            msg = msg
                        };
                        return Json(data1, JsonRequestBehavior.AllowGet);
                    }
                    // Kết thúc check form

                    for (int row = start.Row + 5; row <= end.Row; row++)
                    { // Row by row...
                        var ID = worksheet.Cells[row, 1].Text;
                        msg = "";
                        if (String.IsNullOrEmpty(ID)) break;
                        var isSelected = worksheet.Cells[row, 29].Text?.Trim()?.ToUpper();
                        var findItem = db.InventoryResults.Find(int.Parse(ID));

                        if (findItem == null)
                        {
                            msg = "Không tìm thấy item có ID No : " + ID;
                            goto exitLoop;
                        }
                        findItem.IsSeletedForVerify = isSelected == "Y" ? true : false;
                        if (isSelected == "Y" || isSelected == "YES")
                        {
                            findItem.IsSeletedForVerify = true;
                        }
                        else
                        {
                            findItem.IsSeletedForVerify = false;
                            if (isSelected.Contains("REFER"))
                            {
                                findItem.isReferDie = true;
                            }
                            else
                            {
                                findItem.isReferDie = false;
                            }
                        }

                        db.Entry(findItem).State = EntityState.Modified;
                        db.SaveChanges();
                        selected = isSelected == "Y" ? selected + 1 : selected;


                    exitLoop:
                        ViewBag.NoMean = "Just for exit Loop";
                    }
                    package.SaveAs(output);



                    // 1. Lưu file & control
                    // save file vật lí
                    string fileName = "CAM-Upload-File-Select-Item-" + DateTime.Now.ToString("yyyy-MM-dd-hhmmss");
                    string fileExt = Path.GetExtension(file.FileName);
                    string path = Server.MapPath("~/File/DisposeDie/");
                    fileName += fileExt;
                    package.SaveAs(path + fileName);


                    Dispose_ControlFileUpload newFile = new Dispose_ControlFileUpload()
                    {
                        FileName = fileName,
                        Type = "CAM_Select_Item",
                        Dept = dept,
                        UploadBy = name,
                        UploadDate = DateTime.Now,
                        Active = true
                    };
                    db.Dispose_ControlFileUpload.Add(newFile);
                    db.SaveChanges();

                    output.Position = 0;
                    TempData[handle] = output.ToArray();
                    status = true;
                }
            }

            return Json(new { status = status, selected = selected }, JsonRequestBehavior.AllowGet);
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

