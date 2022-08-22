using ClosedXML.Excel;
using FreelaEdson.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FreelaEdson.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }



        public IActionResult Privacy()
        {
            throw new Exception("Erro na Index");
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public ActionResult ExportLLExcel(DataTableAjaxPostModel model, long mainProjectId = 0)
        {
            model = new DataTableAjaxPostModel();
            model.columns = new List<Column>();
            model.search = new Search();
            model.order = new List<Order>();

            //long userId = ViewBag.UserIdShow;
            long userId = 9999; //HttpContext.Session.Get<long>("UserIdShow");

            ViewBag.MainProjectId = mainProjectId;

            throw new Exception("Teste do ErrorHandlerMeddleware");

            //string officeFilter = mainProjectId == 0 ? model.columns[0].search.value : "";
            //string projectFilter = mainProjectId == 0 ? model.columns[1].search.value : "";
            //string projectNameFilter = mainProjectId == 0 ? model.columns[2].search.value : "";
            //string llCodeFilter = mainProjectId == 0 ? model.columns[3].search.value : model.columns[0].search.value;
            //string equipmentFilter = mainProjectId == 0 ? model.columns[4].search.value : model.columns[1].search.value;
            //string projectPhaseFilter = mainProjectId == 0 ? model.columns[5].search.value : model.columns[2].search.value;
            //string llCategoryFilter = mainProjectId == 0 ? model.columns[6].search.value : model.columns[3].search.value;
            //string problemOpportunityDescriptionFilter = mainProjectId == 0 ? model.columns[7].search.value : model.columns[4].search.value;
            //string businessImpactDescriptionFilter = mainProjectId == 0 ? model.columns[8].search.value : model.columns[5].search.value;
            //string recommendedFutureActionFilter = mainProjectId == 0 ? model.columns[9].search.value : model.columns[6].search.value;
            //string llPriorityFilter = mainProjectId == 0 ? model.columns[10].search.value : model.columns[7].search.value;
            //string actionCommentsFilter = mainProjectId == 0 ? model.columns[11].search.value : model.columns[8].search.value;
            //string y9CodeFilter = mainProjectId == 0 ? model.columns[12].search.value : model.columns[9].search.value;
            //string responsibilityFilter = mainProjectId == 0 ? model.columns[13].search.value : model.columns[10].search.value;
            //string createdByFilter = mainProjectId == 0 ? model.columns[14].search.value : model.columns[11].search.value;
            string createDateFilter = "Teste";// mainProjectId == 0 ? model.columns[15].search.value : model.columns[12].search.value;
            //string llStatusFilter = mainProjectId == 0 ? model.columns[16].search.value : model.columns[13].search.value;

            DateTime createDateD = DateTime.MinValue;

            DateTime.TryParseExact(createDateFilter, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out createDateD);

            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[20] {
                new DataColumn("#", typeof(string)),
                new DataColumn("Office", typeof(string)),
                new DataColumn("Project", typeof(string)),
                new DataColumn("Project Name", typeof(string)),
                new DataColumn("LL Code", typeof(string)),
                new DataColumn("Equipment", typeof(string)),
                new DataColumn("Project Phase",typeof(string)),
                new DataColumn("LL Category",typeof(string)),
                new DataColumn("Problem / Opportunity Description",typeof(string)),
                new DataColumn("Business Impact Description",typeof(string)),
                new DataColumn("Recommended Future Action",typeof(string)),
                new DataColumn("LL Priority",typeof(string)),
                new DataColumn("Action / Comments",typeof(string)),
                new DataColumn("QN Number",typeof(string)),
                new DataColumn("Responsibility",typeof(string)),
                new DataColumn("Created By",typeof(string)),
                new DataColumn("Create Date",typeof(string)),
                new DataColumn("Updated By",typeof(string)),
                new DataColumn("Updated Date",typeof(string)),
                new DataColumn("LL Status",typeof(string))
            });

            //List<ProjectLessonsLearned> result = _db.ProjectLessonsLearneds//.AsNoTracking()
            //   .Where(wh =>
            //        (string.IsNullOrEmpty(officeFilter) || (wh.MainProject.OfficeId.HasValue ? wh.MainProject.ExternalOffice.Name : "NOT DEFINED").ToUpper().Contains(officeFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(projectFilter) || wh.MainProject.SapProjectNumber.ToUpper().Contains(projectFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(projectNameFilter) || wh.MainProject.Description.ToUpper().Contains(projectNameFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(llCodeFilter) || wh.LessonsLearnedCode.ToUpper().Contains(llCodeFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(equipmentFilter) || (wh.EquipmentSizeId.HasValue ? wh.EquipmentSize.EquipmentSizeName : "").ToUpper().Contains(equipmentFilter.ToUpper()) || (wh.CustomDeliverableId.HasValue ? wh.CustomDeliverable.DeliveryName : "").ToUpper().Contains(equipmentFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(projectPhaseFilter) || wh.ProjectPhaseId.ToString() == projectPhaseFilter) &&
            //        (string.IsNullOrEmpty(llCategoryFilter) || wh.LessonsLearnedCategoryId.ToString() == llCategoryFilter) &&
            //        (string.IsNullOrEmpty(problemOpportunityDescriptionFilter) || wh.LessonsLearnedDescription.ToUpper().Contains(problemOpportunityDescriptionFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(businessImpactDescriptionFilter) || wh.LessonsLearnedImpactDescription.ToUpper().Contains(businessImpactDescriptionFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(recommendedFutureActionFilter) || wh.LessonsLearnedRecomFutActionDescription.ToUpper().Contains(recommendedFutureActionFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(llPriorityFilter) || wh.LessonsLearnedPriorityId.ToString() == llPriorityFilter) &&
            //        (string.IsNullOrEmpty(actionCommentsFilter) || wh.LessonsLearnedComments.ToUpper().Contains(actionCommentsFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(y9CodeFilter) || wh.LessonsLearnedY9code.ToUpper().Contains(y9CodeFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(responsibilityFilter) || wh.LessonsLearnedResponsibility.ToUpper().Contains(responsibilityFilter.ToUpper())) &&
            //        (string.IsNullOrEmpty(llStatusFilter) || wh.LessonsLearnedStatusId.ToString() == llStatusFilter) &&
            //        (string.IsNullOrEmpty(createdByFilter) || (wh.UserId.HasValue ? wh.User.UserFullName : "").ToUpper().Contains(createdByFilter.ToUpper())) &&
            //        (createDateD == DateTime.MinValue || ((wh.UpdatedDate ?? wh.CreateDate).Value.Year == createDateD.Year && (wh.UpdatedDate ?? wh.CreateDate).Value.Month == createDateD.Month && (wh.UpdatedDate ?? wh.CreateDate).Value.Day == createDateD.Day)) &&
            //        (mainProjectId == 0 || wh.MainProjectId == mainProjectId)
            //   ).OrderByDescending(a => a.CreateDate).ToList();

            //int count = 1;
            //foreach (ProjectLessonsLearned item in result)
            //{
            //    dt.Rows.Add(
            //        count,
            //        item.MainProject.OfficeId.HasValue ? item.MainProject.ExternalOffice.Name : "NOT DEFINED",
            //        item.MainProject.SapProjectNumber,
            //        item.MainProject.Description,
            //        item.LessonsLearnedCode,
            //        item.EquipmentSizeId.HasValue ? item.EquipmentSize.EquipmentSizeName : "",
            //        item.ProjectPhase.Code,
            //        item.LessonsLearnedCategory.Code,
            //        item.LessonsLearnedDescription,
            //        item.LessonsLearnedImpactDescription,
            //        item.LessonsLearnedRecomFutActionDescription,
            //        item.LessonsLearnedPriority.Code,
            //        item.LessonsLearnedComments,
            //        item.LessonsLearnedY9code,
            //        item.LessonsLearnedResponsibility,
            //        item.UserId.HasValue ? item.User.UserFullName : "",
            //        item.CreateDate.Value.ToString("dd/MMM/yyyy").ToUpper(),
            //        item.UserIdUpdate.HasValue ? item.User.UserFullName : "",
            //        item.UpdatedDate.HasValue ? item.UpdatedDate.Value.ToString("dd/MMM/yyyy").ToUpper() : "",
            //        item.LessonsLearnedStatus.Code
            //    );
            //    count++;
            //}

            //opcao 1
            //NpoiExport nExport = new NpoiExport();
            //nExport.ExportDataTableToWorkbook(dt, "LL");
            //FileContentResult fobject = nExport.ExportDataTableToExcel(dt, "LLExport", "LL");
            //return File(nExport.ExportDataTableToExcel(dt, "LLExport", "LL"), "application/vnd.ms-excel");

            //opcao corrompida
            //FileContentResult fobject;
            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(dt, "LL");
            //    using (MemoryStream stream = new MemoryStream())
            //    {
            //        wb.SaveAs(stream);
            //        FileContentResult bytesdata = File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LLExport.xlsx");
            //        //Content type is used according to latest xlsx format
            //        fobject = bytesdata;
            //    }
            //}
            //return Json(fobject, JsonRequestBehavior.AllowGet);
            //return Json(fobject);

            // Opção do Edson
            //MemoryStream stream = new MemoryStream();
            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(dt, "LL");
            //    wb.SaveAs(stream);
            //}
            ////return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LLExport.xlsx");
            //return new FileContentGeneratingResult("LLExport.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", stream);

            // Outra Opção
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
                }
            }
        }
    }
}
