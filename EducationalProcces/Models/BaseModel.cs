namespace EducationalProcces
{
    public abstract class BaseModel
    {
        public int GetId<T>() => (int)typeof(T).GetProperty($"ID_{typeof(T).Name}")?.GetValue(this);
    }
}