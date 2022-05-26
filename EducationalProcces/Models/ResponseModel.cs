namespace EducationalProcces
{
    public class ResponseModel<T>
    {
        public int StatusCode { get; private set; }
        public string Error { get; private set; }
        public T Data { get; private set; }

        public ResponseModel(int statusCode, string error, T data)
        {
            StatusCode = statusCode;
            Error = error;
            Data = data;
        }
    }
}
