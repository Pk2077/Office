using System;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace OfficeManagement.Models
{
    public class Employee
    {
        private static string _connectionString = ConfigurationManager.ConnectionStrings["default"].ConnectionString;
        public int Id { get; set; }
        public string Name { get; set; }
        public string Designation { get; set; }
        public DateTime DateOfJoin { get; set; }
        public decimal Salary { get; set; }
        public string Gender { get; set; }
        public string State { get; set; }
        public List<Employee> GetEmployees()
        {
            string sQuery = "select * from Employees";
            List<Employee> _Employees = new List<Employee>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var command = new SqlCommand(sQuery, connection))
                {
                    var reader = command.ExecuteReader();
                    while (reader.Read()) {
                        Employee employee = new Employee();
                        employee.Id = reader.GetInt32(reader.GetOrdinal("Id"));
                        employee.Name = reader.GetString(reader.GetOrdinal("Name"));
                        employee.Designation = reader.GetString(reader.GetOrdinal("Designation"));
                        employee.DateOfJoin = reader.GetDateTime(reader.GetOrdinal("DateOfJoin"));
                        employee.Salary = reader.GetDecimal(reader.GetOrdinal("Salary"));
                        employee.Gender = reader.GetString(reader.GetOrdinal("Gender"));
                        employee.State = reader.GetString(reader.GetOrdinal("State"));
                        _Employees.Add(employee);
                    }
                }
            }
            //Session["Employees"] = _Employees;
            return _Employees;
        }
        public void SaveEmployee(Employee employee)
        {
            string sQuery = string.Empty;
            if (employee.Id > 0)
            {
                sQuery = $@"update Employees set Name='{employee.Name}',Designation='{employee.Designation}',DateOfJoin='{employee.DateOfJoin.ToString("MM-dd-yyyy")}'
                          ,Salary={employee.Salary},Gender='{employee.Gender}',State='{employee.State}' 
                           where Id = {employee.Id};";
            }
            else
            {
               sQuery = @"Insert into Employees(Name,Designation,DateOfJoin,Salary,Gender,State) 
                  Values(@Name, @Designation, @DateOfJoin, @Salary, @Gender, @State)";
               
            }
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (SqlCommand cmd = new SqlCommand(sQuery, connection))
                {
                    cmd.Parameters.AddWithValue("@Name", employee.Name);
                    cmd.Parameters.AddWithValue("@Designation", employee.Designation);
                    cmd.Parameters.AddWithValue("@DateOfJoin", employee.DateOfJoin);
                    cmd.Parameters.AddWithValue("@Salary", employee.Salary);
                    cmd.Parameters.AddWithValue("@Gender", employee.Gender);
                    cmd.Parameters.AddWithValue("@State", employee.State);

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public List<Employee> DeleteEmployee(int id)
        {
            string sQuery = "delete from Employees where Id="+id;
            List<Employee> _Employees = new List<Employee>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var command = new SqlCommand(sQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
            return _Employees;
        }
    }
    public class State
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}