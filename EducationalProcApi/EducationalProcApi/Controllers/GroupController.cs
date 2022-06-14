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
    public class GroupController : ControllerBase
    {
        /// <summary>
        /// Метод получения данных
        /// </summary>
        /// <param name="whereColumnName">переменная для передачи данных об имени столбца</param>
        /// <param name="whereValue">переменная для передачи данных об значении</param>
        /// <param name="orderByColumnName">переменная для передачи порядка по имени столбца</param>
        /// <returns></returns>
        [HttpGet("{whereColumnName}/{whereValue}/{orderByColumnName}")]
        public async Task<ActionResult<Group>> GetGroup(string whereColumnName, string whereValue, string orderByColumnName)
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
                    List<Group> result = db.ResultModels.FromSqlRaw("EXEC dbo.GetGroups @WhereColumnName, @WhereValue, @OrderByColumnName", parameters).ToList().Select(r => r.ParseResult<Group>()).ToList();

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
        /// Мето изменения данных о группе
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_Group</param>
        /// <param name="Group">переменная для обращения к модели Group</param>
        /// <returns></returns>
        [HttpPut("{id}")]
        public async Task<IActionResult> PutGroup(int id, Group Group)
        {
            if (id != Group.ID_Group)
                return BadRequest();

            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Group", SqlDbType.Int).WithValue(Group.ID_Group),
                    new SqlParameter("@Name_Group", SqlDbType.VarChar, 30).WithValue(Group.Name_Group),
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Group_update @ID_Group, @Name_Group", parameters);
                    SqlParameter parameter = new SqlParameter("@ID_Group", SqlDbType.Int).WithValue(Group.ID_Group);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetGroup @ID_Group", parameter).ToList()[0].ParseResult<Group>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод добавлния данных о группе
        /// </summary>
        /// <param name="Group">пременная для обращения к модели Group</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<ActionResult<Group>> PostGroup(Group Group)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {

                    new SqlParameter("@Name_Group", SqlDbType.VarChar, 30).WithValue(Group.Name_Group),
                    new SqlParameter("@ID_Group", SqlDbType.Int)
                    {
                       Direction = ParameterDirection.Output
                    }
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Group_insert @Name_Group, @ID_Group OUT", parameters);

                    SqlParameter parameter = new SqlParameter("@ID_Group", SqlDbType.Int).WithValue((int)parameters[6].Value);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetGroup @ID_Group", parameter).ToList()[0].ParseResult<Group>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод удаления группы
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_Group</param>
        /// <returns></returns>
        [HttpDelete("{id}")] //+
        public async Task<IActionResult> DeleteGroup(int id)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_Group", SqlDbType.Int).WithValue(id)
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Group_delete @ID_Group", parameters);
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