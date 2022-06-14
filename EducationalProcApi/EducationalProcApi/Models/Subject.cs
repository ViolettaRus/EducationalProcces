using System.ComponentModel.DataAnnotations;

namespace EducationalProc
{
    public class Subject
    {
        [Key]
        public int ID_Subject { get; set; }

        private string _name_subject;

        public string Name_Subject
        {
            get => _name_subject;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 30))
                    _name_subject = value;

            }
        }
    }
}