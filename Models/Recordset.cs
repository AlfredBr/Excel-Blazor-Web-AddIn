namespace excel_blazor_web_addin.Models;

public class Recordset<T>
{
    public Recordset(IEnumerable<T>? records)
    {
        Headers = GetHeaders(records);
        Records = records ?? new List<T>();
    }
    private IEnumerable<string> GetHeaders(IEnumerable<T>? records)
    {
        if (records == null)
        {
            return new List<string>();
        }

        var type = typeof(T);
        var properties = type.GetProperties();
        var headers = properties.Select(p => p.Name).ToList();
        return headers;
    }
    public IEnumerable<string> Headers { get; set; }
    public IEnumerable<T> Records { get; set; }
}
