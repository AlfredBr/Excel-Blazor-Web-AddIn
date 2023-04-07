namespace excel_blazor_web_addin.Models;

public interface IViewModel
{
    bool IsReady { get; }
    public bool IsLoading { get; }
}
