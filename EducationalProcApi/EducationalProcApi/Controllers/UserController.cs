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
    public class UserController : ControllerBase
    {
        /// <summary>
        /// Метод получения данных
        /// </summary>
        /// <param name="whereColumnName">переменная для передачи данных об имени столбца</param>
        /// <param name="whereValue">переменная для передачи данных об значении</param>
        /// <param name="orderByColumnName">переменная для передачи порядка по имени столбца</param>
        /// <returns></returns>
        [HttpGet("{whereColumnName}/{whereValue}/{orderByColumnName}")] //+ 
        public async Task<ActionResult<User>> GetUsers(string whereColumnName, string whereValue, string orderByColumnName)
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
                    List<User> result = db.ResultModels.FromSqlRaw("EXEC dbo.GetUsers @WhereColumnName, @WhereValue, @OrderByColumnName", parameters).ToList().Select(r => r.ParseResult<User>()).ToList();

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
        /// <summary>
        /// Метод получения данных 
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_User</param>
        /// <returns></returns>
        [HttpGet("{id}")] //+
        public async Task<ActionResult<User>> GetUser(int id)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using DataBaseHelper db = new();
                    SqlParameter parameter = new SqlParameter("@ID_User", SqlDbType.Int).WithValue(id);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetUser @ID_User", parameter).ToList()[0].ParseResult<User>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод проверки логина пользователя
        /// </summary>
        /// <param name="login">переменная для передачи данных логина</param>
        /// <param name="password">переменная для передачи данных пароля</param>
        /// <returns></returns>
        [HttpGet("{login}/{password}")] 
        public async Task<ActionResult<User>> GetLoggedUser(string login, string password)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using DataBaseHelper db = new DataBaseHelper();
                    List<User> users = db.Users.Where(u => (u.Login == login) && (u.Password == password)).ToList();

                    if (users.Count > 0)
                    {
                        SqlParameter parameter = new SqlParameter("@ID_User", SqlDbType.Int).WithValue(users[0].ID_User);
                        return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetUser @ID_User", parameter).ToList()[0].ParseResult<User>());
                    }
                    else
                        throw new Exception("Пользователь не найден");
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод изменения данных о пользователе
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_User</param>
        /// <param name="users">переменная для обращения к модели  User</param>
        /// <returns></returns>
        [HttpPut("{id}")] 
        public async Task<IActionResult> PutUser(int id, User users)
        {
            if (id != users.ID_User)
                return BadRequest();

            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_User", SqlDbType.Int).WithValue(users.ID_User),
                    new SqlParameter("@Login", SqlDbType.VarChar, 100).WithValue(users.Login),
                    new SqlParameter("@Password", SqlDbType.VarChar, 200).WithValue(users.Password),
                    new SqlParameter("@Role_ID", SqlDbType.VarChar, 200).WithValue(users.Role.ID_Role),

                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Users_update @ID_User, @Login, @Password, @Role_ID", parameters);
                    SqlParameter parameter = new SqlParameter("@ID_User", SqlDbType.Int).WithValue(users.ID_User);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetUser @ID_User", parameter).ToList()[0].ParseResult<User>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод добавления данных о пользователе
        /// </summary>
        /// <param name="users">пременная для обращения к моделе User</param>
        /// <returns></returns>
        [HttpPost] 
        public async Task<ActionResult<User>> PostUser(User users)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {                   
                    new SqlParameter("@Login", SqlDbType.VarChar, 100).WithValue(users.Login),
                    new SqlParameter("@Password", SqlDbType.VarChar, 200).WithValue(users.Password),
                    new SqlParameter("@Role_ID", SqlDbType.VarChar, 200).WithValue(users.Role.ID_Role),
                    new SqlParameter("@ID_User", SqlDbType.Int)
                    {
                       Direction = ParameterDirection.Output
                    }
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Users_insert @Login, @Password, @Role_ID, @ID_User OUT", parameters);

                    SqlParameter parameter = new SqlParameter("@ID_User", SqlDbType.Int).WithValue((int)parameters[3].Value);
                    return StatusCode(201, db.ResultModels.FromSqlRaw("EXEC dbo.GetUser @ID_User", parameter).ToList()[0].ParseResult<User>());
                }
                catch (Exception ex)
                {
                    return StatusCode(400, ex.Message);
                }
            });
        }
        /// <summary>
        /// Метод удаления данных о пользователе
        /// </summary>
        /// <param name="id">переменная для передачи данных из ID_User</param>
        /// <returns></returns>
        [HttpDelete("{id}")] //+
        public async Task<IActionResult> DeleteUser(int id)
        {
            return await Task.Run(() =>
            {
                using DataBaseHelper db = new();
                SqlParameter[] parameters =
                {
                    new SqlParameter("@ID_User", SqlDbType.Int).WithValue(id)
                };
                try
                {
                    db.Database.ExecuteSqlRaw("EXEC dbo.Users_delete @ID_User", parameters);
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