using EPPlus.Core.Extensions;
using OfficeOpenXml;
using System;
using System.Reflection;

namespace WebUploadExcel
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (!FileUpload1.HasFile) return;
            using (var excelPackage = new ExcelPackage(FileUpload1.FileContent))
            {
                var dtos = excelPackage.ToList<PersonDto>(1, configuration => configuration.SkipCastingErrors());
                Response.ContentType = "application/json";
                Response.Write(Newtonsoft.Json.JsonConvert.SerializeObject(dtos));
                Response.End();
            }
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            using (var excelPackage = Assembly.GetExecutingAssembly().GenerateExcelPackage(nameof(PersonDto)))
            {
                Response.AppendHeader("Content-Disposition", $"attachment; filename={nameof(PersonDto)}.xls");
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                excelPackage.SaveAs(Response.OutputStream);
                Response.End();
            }
        }
    }
}