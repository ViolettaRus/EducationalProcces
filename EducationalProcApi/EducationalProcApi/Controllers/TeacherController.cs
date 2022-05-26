using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace EducationalProc
{
    [Route("api/[controller]")]
    [ApiController]
    public class TeacherController : ControllerBase
    {
        [HttpGet("{whereColumnName}/{whereValue}/{orderByColumnName}")] 
        public async Task<ActionResult<Teacher>> GetTeacher(string whereColumnName, string whereValue, string orderByColumnName)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using DataBaseHelper db = new();
                    SqlParameter[] parameters =
                    {
                        new SqlParameter("@WhereColumnName", SqlDbType.VarChar, 200).WithValue(whereColumnName),
                        new SqlParameter("@WhereValue", SqlDbType.VarChar, 200).WithValue(whereValue),
                        new SqlParameter("@OrderByColumnName", SqlDbType.VarChar, 200).WithValue(orderByColumnName)
                    };
                    List<Teacher> result = db.ResultModels.FromSqlRaw("EXEC dbo.GetTeachers @WhereColumnName, @WhereValue, @OrderByColumnName", parameters).ToList().Select(r => r.ParseResult<Teacher>()).ToList();

                    if (result.Count == 0)
                        throw new Exception("Пользователь не найден");
                    else
                        return StatusCode(201, result);
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> PutTeacher(int id, Teacher Teacher)
        {
            if (id != Teacher.ID_Teacher)
                return BadRequest();

            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Teacher", SqlDbType.Int).WithValue(Teacher.ID_Teacher),
                    new SqlParameter("@FIO", SqlDbType.VarChar, 200).WithValue(Teacher.FIO),
                    new SqlParameter("@Phone", SqlDbType.VarChar, 15).WithValue(Teacher.Phone),
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Teacher_update @ID_Teacher, @FIO, @Phone", parameters);
                    SqlParameter parameter = new SqlParameter("@ID_Teacher", SqlDbType.Int).WithValue(Teacher.ID_Teacher);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetTeacher @ID_Teacher", parameter).ToList()[0].ParseResult<Teacher>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }

        [HttpPost]
        public async Task<ActionResult<Teacher>> PostTeacher(Teacher Teacher)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {

                    new SqlParameter("@FIO", SqlDbType.VarChar, 200).WithValue(Teacher.FIO),
                    new SqlParameter("@Phone", SqlDbType.VarChar, 15).WithValue(Teacher.Phone),
                    new SqlParameter("@ID_Teacher", SqlDbType.Int)
                    {
                       Direction = ParameterDirection.Output
                    }
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Teacher_insert @FIO, @Phone, @ID_Teacher OUT", parameters);

                    SqlParameter parameter = new SqlParameter("@ID_Teacher", SqlDbType.Int).WithValue((int)parameters[5].Value);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetTeacher @ID_Teacher", parameter).ToList()[0].ParseResult<Teacher>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }

        [HttpDelete("{id}")] //+
        public async Task<IActionResult> DeleteTeacher(int id)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Teacher", SqlDbType.Int).WithValue(id)
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Teacher_delete @ID_Teacher", parameters);
                    return StatusCode(204, null);
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }

    }
}
