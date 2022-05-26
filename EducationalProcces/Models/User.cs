using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace EducationalProcces
{
    public class User : BaseModel
    {
        [Key]
        public int ID_User { get; set; }

        private string _login;

        public string Login
        {
            get => _login;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 50))
                    _login = value;

            }
        }

        private string _password;

        public string Password
        {
            get => _password;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 50))
                    _password = value;

            }
        }

        [ForeignKey("Role_ID")]
        public Role Role { get; set; } = new Role();



    }
}
