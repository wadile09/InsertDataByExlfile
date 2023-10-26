using InsertDataByExcelFile.Models;
using Microsoft.AspNetCore.Mvc;
using sendEmail.codeFunction;
using System.Data;
using System.Data.SqlClient;

namespace InsertDataByExcelFile.Controllers
{
	public class StudentController : Controller
    {


		private readonly genarat_Code _genaratCode;

		public IConfiguration Configuration { get; }

		public StudentController(IConfiguration configuration)
		{
			Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
			_genaratCode = new genarat_Code(configuration);
		}


        public IActionResult Index()
        {
            var list = new List<I_StudentModel>();
            string connectionString = Configuration.GetConnectionString("myDbConnection");
            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand("SP_Student_Data", sqlConnection))
                {

                    command.CommandType = CommandType.StoredProcedure;
                    sqlConnection.Open();
                    using (SqlDataReader sdr = command.ExecuteReader())
                    {
                        while (sdr.Read())
                        {
                            var student = new I_StudentModel()
                            {
                                Id = (int)sdr["Id"],
                                Name = (string)sdr["StudentName"],
                                Email = (string)sdr["Email"],
                                RollNo = (string)sdr["RollNo"]
                            };
                            list.Add(student);
                        }
                    }
                }
            }
            return View(list);
        }

        public IActionResult Post(IFormFile file)
        {
			_genaratCode.Generatelist(file);

	   //     var list= new List<StudentModel>();
	   //     using (var stream = new MemoryStream())
	   //     {
	   //         file.CopyToAsync(stream);
	   //         using (var package = new ExcelPackage(stream))
	   //         {

			//             ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
			//             var rowcount = worksheet.Dimension.Rows;
			//             for (int row = 2; row <= rowcount; row++)
			//             {
			//                 list.Add( new StudentModel
			//                 {
			//                     Name = worksheet.Cells[row, 1]?.Value?.ToString()?.Trim() ?? string.Empty,
			//                     RollNo = worksheet.Cells[row, 2]?.Value?.ToString()?.Trim() ?? string.Empty,
			//                     Email = worksheet.Cells[row, 3]?.Value?.ToString()?.Trim() ?? string.Empty,

			//                 });

			//                 string connectionString = Configuration.GetConnectionString("myDbConnection");
			//                 SqlConnection sqlConnection = new SqlConnection(connectionString);
			//                 SqlCommand command = new SqlCommand();
			//                 SqlDataAdapter adp = new SqlDataAdapter(command);
			//                 sqlConnection.Open();
			//                 foreach (var item in list)
			//                 {
			//                     command.CommandText = "Checkdata";
			//                     command.CommandType = CommandType.StoredProcedure;
			//                     command.Connection = sqlConnection;
			//                     var param = command.CreateParameter();
			//                     param.ParameterName = "@RollNo";
			//                     param.Value = item.RollNo;
			//                     command.Parameters.Add(param);

			//                 }
			//                 int result = Convert.ToInt32(command.ExecuteScalar());
			//                 sqlConnection.Close();
			//                 foreach (var item in list)
			//                 {
			//                     sqlConnection.Open();
			//                    var flag = string.Empty;
			//                     if (result == 1)
			//                     {
			//                         flag = "U";

			//                     }
			//                     command.Parameters.Clear();
			//                     command.CommandText = "InsertStudent";
			//                     command.CommandType = CommandType.StoredProcedure;
			//                     command.Connection = sqlConnection;
			//                     command.Parameters.AddWithValue("@Name", item.Name);
			//                     command.Parameters.AddWithValue("@RollNo", item.RollNo);
			//                     command.Parameters.AddWithValue("@Email", item.Email);
			//                     command.Parameters.AddWithValue("@Flag", flag);
			//                     command.ExecuteNonQuery();

			//email.SendMail(item.Email);
			//                 }
			//                 TempData["Done"] = "Successfully data Insert";

			//                 list.Clear();
			//             }
			//         };
			//     }

			return RedirectToAction("Index");
        }
        
    }

}

