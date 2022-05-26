using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace EducationalProcces
{
    public class Group : BaseModel
    {
        [Key]
        public int ID_Group { get; set; }

        private string _name_group;

        public string Name_Group
        {
            get => _name_group;
            set
            {
                if ((!string.IsNullOrWhiteSpace(value)) && (value.Length <= 30))
                    _name_group = value;

            }
        }
    }
}
