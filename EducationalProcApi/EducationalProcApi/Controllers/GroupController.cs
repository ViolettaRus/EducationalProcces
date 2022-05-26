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
