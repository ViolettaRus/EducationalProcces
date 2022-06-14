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
    public class SubjectController : ControllerBase
    {
        /// <summary>
        /// Метод получения данных
        /// </summary>
        /// <param name="whereColumnName">переменная для передачи данных об имени столбца</param>
        /// <param name="whereValue">переменная для передачи данных об значении</param>
        /// <param name="orderByColumnName">переменная для передачи порядка по имени столбца</param>
        /// <returns></returns>
        [HttpGet("{whereColumnName}/{whereValue}/{orderByColumnName}")]
        public async Task<ActionResult<Subject>> GetSubject(string whereColumnName, string whereValue, string orderByColumnName)
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
                    List<Subject> result = db.ResultModels.FromSqlRaw("EXEC dbo.GetSubjects @WhereColumnName, @WhereValue, @OrderByColumnName", parameters).ToList().Select(r => r.ParseResult<Subject>()).ToList();

                    if (result.Count == 0)
                        throw new Exception("Группа не найдена");
                    else
                        return StatusCode(201, result);
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод изменения данных о предметах
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_Subject</param>
        /// <param name="Subject">пременная для обращения к модулю Subject</param>
        /// <returns></returns>
        [HttpPut("{id}")]
        public async Task<IActionResult> PutSubject(int id, Subject Subject)
        {
            if (id != Subject.ID_Subject)
                return BadRequest();

            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Subject", SqlDbType.Int).WithValue(Subject.ID_Subject),
                    new SqlParameter("@Name_Subject", SqlDbType.VarChar, 30).WithValue(Subject.Name_Subject),
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Subject_update @ID_Subject, @Name_Subject", parameters);
                    SqlParameter parameter = new SqlParameter("@ID_Subject", SqlDbType.Int).WithValue(Subject.ID_Subject);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetSubject @ID_Subject", parameter).ToList()[0].ParseResult<Subject>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод добавления данных о предметах
        /// </summary>
        /// <param name="Subject">пременная для обращения к модулю Subject</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<ActionResult<Subject>> PostSubject(Subject Subject)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {

                    new SqlParameter("@Name_Subject", SqlDbType.VarChar, 30).WithValue(Subject.Name_Subject),
                    new SqlParameter("@ID_Subject", SqlDbType.Int)
                    {
                       Direction = ParameterDirection.Output
                    }
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Subject_insert @Name_Subject, @ID_Subject OUT", parameters);

                    SqlParameter parameter = new SqlParameter("@ID_Subject", SqlDbType.Int).WithValue((int)parameters[6].Value);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetSubject @ID_Subject", parameter).ToList()[0].ParseResult<Subject>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод удаления данных о предметах
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_Subject</param>
        /// <returns></returns>
        [HttpDelete("{id}")] //+
        public async Task<IActionResult> DeleteSubject(int id)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Subject", SqlDbType.Int).WithValue(id)
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Subject_delete @ID_Subject", parameters);
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