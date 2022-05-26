using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EducationalProcces
{
    public abstract class BaseModel
    {
        public int GetId<T>() => (int)typeof(T).GetProperty($"ID_{typeof(T).Name}")?.GetValue(this);
    }
}
