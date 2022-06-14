using System.ComponentModel.DataAnnotations;

namespace EducationalProcces
{
    public class Role : BaseModel
    {
        [Key]
        public int ID_Role { get; set; }

        private string _role;

        public string Name_Role
        {
            get => _role;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 200))
                    _role = value;

            }
        }
    }
}