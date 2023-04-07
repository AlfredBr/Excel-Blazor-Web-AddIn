using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;
using System.Drawing;
using System.Reflection;

namespace excel_blazor_web_addin.Pages;

public partial class Index : IDisposable
{
    [Inject] public IJSRuntime JSRuntime { get; set; } = default!;
    public IJSObjectReference JSModule { get; set; } = default!;
    private string OfficeContext { get; set; } = string.Empty;
    private Size? _size;
    private DotNetObjectReference<Index>? _objRef;
    private bool _disposedValue = false;
    protected override async Task OnInitializedAsync()
    {
        // TODO: Add your initialization logic here
        await base.OnInitializedAsync();
    }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            Debug.Assert(JSRuntime is not null);
            _objRef = DotNetObjectReference.Create(this);
            Debug.Assert(_objRef is not null);
            _ = await JSRuntime.InvokeAsync<string>("SetDotNetHelper", _objRef);
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");
            Debug.Assert(JSModule is not null);
            OfficeContext = await JSModule.InvokeAsync<string>("getOfficeContextAsync");
            await InvokeAsync(StateHasChanged);
            await base.OnAfterRenderAsync(firstRender);
        }
    }

    [JSInvokable]
    public async Task OnResize(int width, int height)
    {
        Console.WriteLine($"OnResize({width},{height})");
        _size = new Size(width, height);
        await InvokeAsync(StateHasChanged);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // dispose managed state (managed objects)
                _objRef?.Dispose();
            }

            // free unmanaged resources (unmanaged objects) and override finalizer and set large fields to null
            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
