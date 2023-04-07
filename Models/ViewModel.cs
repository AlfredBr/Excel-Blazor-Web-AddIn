namespace excel_blazor_web_addin.Models;

public class ViewModel<T> : IViewModel
{
    public bool IsReady => Items?.Any() ?? false;
    public bool IsLoading => !IsReady;
    public IEnumerable<T>? Items { get; set; }
}
