using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace EducationalProcces
{
    public static class APIHelper
    {
        private static readonly HttpClient _client = new();
        private static string _link { get; set; }
        /// <summary>
        /// Метод обращения к json файлу
        /// </summary>
        public static void ActivateLink()
        {
            var builder = new ConfigurationBuilder();
            builder.SetBasePath(Directory.GetCurrentDirectory());
            builder.AddJsonFile("appsettings.json");
            var config = builder.Build();
            _link = config.GetSection("ConnectionLink").Value;
        }
        /// <summary>
        /// Метод добавления данных
        /// </summary>
        /// <typeparam name="T">переменная Т</typeparam>
        /// <param name="data">переменная date</param>
        /// <returns></returns>
        public static async Task<ResponseModel<T>> PostDataAsync<T>(T data)
        {
            var stringContent = new StringContent(JsonSerializer.Serialize(data), Encoding.UTF8, "application/json");
            return await GetResponseAsync<T>(await _client.PostAsync($"{_link}{data.GetType().Name}", stringContent));
        }
        /// <summary>
        /// Метот изменения данных
        /// </summary>
        /// <typeparam name="T">переменная Т</typeparam>
        /// <param name="data">переменная date</param>
        /// <returns></returns>
        public static async Task<ResponseModel<T>> PutDataAsync<T>(T data)
        {
            var stringContent = new StringContent(JsonSerializer.Serialize(data), Encoding.UTF8, "application/json");
            return await GetResponseAsync<T>(await _client.PutAsync($"{_link}{data.GetType().Name}/{(data as BaseModel).GetId<T>()}", stringContent));
        }
        /// <summary>
        /// Метод удаления данных
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static async Task<ResponseModel<T>> DeleteDataAsync<T>(T data) =>
            await GetResponseAsync<T>(await _client.DeleteAsync($"{_link}{data.GetType().Name}/{(data as BaseModel).GetId<T>()}"));
        /// <summary>
        /// Метод передачи данных
        /// </summary>
        /// <typeparam name="T">переменная date Т</typeparam>
        /// <param name="id">переменная для передачи данных через ID</param>
        /// <returns></returns>
        public static async Task<ResponseModel<T>> GetDataAsync<T>(int id) =>
            await GetResponseAsync<T>(await _client.GetAsync($"{_link}{typeof(T).Name}/{id}"));
        /// <summary>
        /// Метод получения данных
        /// </summary>
        /// <typeparam name="T">переменная Т</typeparam>
        /// <param name="whereColumnName">переменная для передачи данных об имени столбца</param>
        /// <param name="whereValue">переменная для передачи данных об значении</param>
        /// <param name="orderByColumnName">переменная для передачи порядка по имени столбца</param>
        /// <param name="dataType">переменная dateType</param>
        /// <returns></returns>
        public static async Task<ResponseModel<T>> GetDataAsync<T>(string whereColumnName, string whereValue, string orderByColumnName, Type dataType) =>
            await GetResponseAsync<T>(await _client.GetAsync($"{_link}{dataType.Name}/{whereColumnName}/{whereValue}/{orderByColumnName}"));
        /// <summary>
        /// Метод проверки логина и пароля пользователя
        /// </summary>
        /// <param name="login">переменная для передачи данных через login</param>
        /// <param name="password">переменная для передачи данных через password</param>
        /// <returns></returns>
        public static async Task<ResponseModel<User>> GetLoggedUser(string login, string password) => await GetResponseAsync<User>(await _client.GetAsync($"{_link}Users/{login}/{password}"));
        /// <summary>
        /// Метод ответа данных
        /// </summary>
        /// <typeparam name="T">переменная T</typeparam>
        /// <param name="response">переменная ответа</param>
        /// <returns></returns>
        private static async Task<ResponseModel<T>> GetResponseAsync<T>(HttpResponseMessage response)
        {
            try
            {
                response.EnsureSuccessStatusCode();

                switch ((int)response.StatusCode)
                {
                    case 201:
                        {
                            return new ResponseModel<T>((int)response.StatusCode, null, JsonSerializer.Deserialize<T>(await response.Content.ReadAsStringAsync(), new JsonSerializerOptions { PropertyNameCaseInsensitive = true }));
                        }
                    case 204:
                        {
                            return new ResponseModel<T>((int)response.StatusCode, null, default);
                        }
                    default:
                        {
                            return new ResponseModel<T>(0, null, default);
                        }
                }
            }
            catch (HttpRequestException ex)
            {
                return new ResponseModel<T>((int)response.StatusCode, await response.Content.ReadAsStringAsync(), default);
            }
        }
    }
}