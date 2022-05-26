using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace EducationalProcces
{
    public class Teacher : BaseModel
    {
        [Key]
        public int ID_Teacher { get; set; }

        private string _fio;
        public string FIO
        {
            get => _fio;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 200))
                    _fio = value;

            }
        }

        private string _phone;
        public string Phone
        {
            get => _phone;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 16))
                    _phone = value;
            }
        }
    }
}
