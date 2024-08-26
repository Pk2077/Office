using ClosedXML.Excel;
using OfficeManagement.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficeManagement.Controllers
{
    public class EmployeeController : Controller
    {
        public ActionResult Index()
        {
            Employee _Employee = new Employee();
            List<Employee> list = _Employee.GetEmployees();
            Session["Employees"] = list;
            return View(list);
        }
        public ActionResult GetEmployees()
        {
            Employee _Employee = new Employee();
            List<Employee> list = _Employee.GetEmployees();
            Session["Employees"] = list;
            return PartialView("List", list);
        }
        public ActionResult NewEmployee()
        {
            Employee _Employee = new Employee();
            _Employee.DateOfJoin = DateTime.Now;
            return PartialView(_Employee);
        }
        public ActionResult EditEmployee(int id)
        {
            Employee _Employee = new Employee();
            var emplist = _Employee.GetEmployees();
            if (emplist.Count > 0 && emplist.Where(x => x.Id == id).ToList().Count > 0)
            {
                _Employee = emplist.Where(x => x.Id == id).SingleOrDefault();
            }
            return PartialView("NewEmployee", _Employee);
        }
        public ActionResult SaveEmployee(int Id, string Name, string Designation, string DateOfJoin, string Salary, string Gender, string State)
        {
            try
            {
                Employee _Employee = new Employee();
                _Employee.Id = Id;
                _Employee.Name = Name;
                _Employee.Designation = Designation;
                _Employee.DateOfJoin = Convert.ToDateTime(DateOfJoin);
                _Employee.Salary = Convert.ToDecimal(Salary);
                _Employee.Gender = Gender;
                _Employee.State = State;
                _Employee.SaveEmployee(_Employee);
                if (_Employee.DateOfJoin > DateTime.Now)
                {
                    return Json(new { status = false, Message = "You can't select a future date." });
                }
                return Json(new { status = true, Message = "Success" });
            }
            catch (Exception ex)
            {
                return Json(new { status = false, Message = "Failed : " + ex.Message });
            }
        }
        public ActionResult DeleteEmployee(int id)
        {
            try
            {
                if (id > 0)
                {
                    new Employee().DeleteEmployee(id);
                    return Json(new { status = true, Message = "Succesfully deleted." });
                }
                else
                {
                    return Json(new { status = true, Message = "Invalid employee id." });
                }
            }
            catch (Exception ex)
            {
                return Json(new { status = false, Message = "Failed to delete : " + ex.Message });
            }
        }
        public void ExportExcel()
        {
            List<Employee> Employeelst = new List<Employee>();
            if (Session["Employees"] != null)
            {
                Employeelst = (List<Employee>)Session["Employees"];
            }
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("S.No", typeof(int));
            dt.Columns.Add("Name");
            dt.Columns.Add("Designation");
            dt.Columns.Add("DOJ");
            dt.Columns.Add("Salary",typeof(decimal));
            dt.Columns.Add("Gender");
            dt.Columns.Add("State");
            int i = 1;
            foreach (Employee row in Employeelst)
            {
                DataRow dRow = dt.NewRow();
                dRow["S.No"] = i;
                dRow["Name"] = row.Name;
                dRow["Designation"] = row.Designation;
                dRow["DOJ"] = row.DateOfJoin;
                dRow["Salary"] = row.Salary;
                dRow["Gender"] = row.Gender;
                dRow["State"] = row.State;
                dt.Rows.Add(dRow);
                i++;
            }
            WorkbookXML(dt, "Employees Report", "A1:G1", "A2:G2", "A3:G3");
        }
        public void WorkbookXML(DataTable dt, string Head, string Titles1, string Titles2, string Titles3)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                // Ensure Titles are within valid Excel ranges
                ws.Range(Titles1).Merge().AddToNamed("Titles");
                ws.Range(Titles1).Value = "PK-App-1.0";
                ws.Range(Titles2).Merge().AddToNamed("Titles");
                ws.Range(Titles2).Value = Head;
                ws.Range(Titles3).Merge().AddToNamed("Titles");
                ws.Range(Titles3).Value = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");

                if (Head == "Employees Report")
                {
                    ws.Column("A").Width = 5;
                    ws.Column("B").Width = 22;
                    ws.Column("C").Width = 22;
                    ws.Column("D").Width = 22;
                    ws.Column("E").Width = 22;
                    ws.Column("F").Width = 22;
                    ws.Column("G").Width = 15;
                }

                var tableWithData = ws.Cell(5, 1).InsertTable(dt.AsEnumerable());
                ws.Tables.FirstOrDefault().ShowAutoFilter = false; // To remove Filter in Table Heading

                // Style adjustments
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                wb.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                wb.Style.Font.Bold = true;

                var titlesStyle = wb.Style;
                titlesStyle.Font.Bold = true;
                titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                titlesStyle.Fill.BackgroundColor = XLColor.Olive;
                wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                wb.Style.Font.Bold = true;

                // Send Excel file to client
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + Head.Replace(" ", "") + "_" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + ".xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }

    }
}