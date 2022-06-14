using Microsoft.EntityFrameworkCore;

namespace EducationalProc
{
    public class DataBaseHelper : DbContext
    { 
        public DbSet<Role> Role { get; set; }
        public DbSet<User> Users { get; set; }
        public DbSet<Teacher> Teacher { get; set; }
        public DbSet<Group> Group { get; set; }
        public DbSet<Subject> Subject { get; set; }
        public DbSet<ResultModel> ResultModels { get; set; }


        public DataBaseHelper(DbContextOptions<DataBaseHelper> options)
            : base(options)
        {

        }

        public DataBaseHelper()
        {

        }
        /// <summary>
        /// Метод полключения к локальной базе данных
        /// </summary>
        /// <param name="optionsBuilder"></param>
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder) =>
            optionsBuilder.UseSqlServer("Data Source=DESKTOP-9F8J8EP\\SQLRUS;Initial Catalog=EducationalProcces;User ID=sa;Password=12345678;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ResultModel>(entity =>
            {
                entity.HasNoKey();
            });            
        }
    }
}