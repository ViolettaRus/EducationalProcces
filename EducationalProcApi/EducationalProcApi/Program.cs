using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;

namespace EducationalProc
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }
        /// <summary>
        /// Метод обращения к классу Startup
        /// </summary>
        /// <param name="args">переменная</param>
        /// <returns></returns>
        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
    }
}