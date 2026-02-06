using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

internal interface IRibbonTick
{
    Task Tick_50ms(int? progresso = null);
}

internal sealed class RibbonTickComProgresso : IRibbonTick
{
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
    private long _ultimoTick;
    private readonly IProgress<int> _progress;

    public RibbonTickComProgresso(IProgress<int> progress)
    {
        _progress = progress;
        _ultimoTick = 0;
    }

    public async Task Tick_50ms(int? progresso = null)
    {
        // limite de frequência (50ms)
        if (_stopwatch.ElapsedMilliseconds - _ultimoTick >= 50)
        {
            _ultimoTick = _stopwatch.ElapsedMilliseconds;

            if (progresso.HasValue)
                _progress.Report(progresso.Value);

            // libera UI do Word (STA safe)
            await Task.Yield();
        }
    }
}
internal sealed class RibbonTickNenhum : IRibbonTick
{
    public Task Tick_50ms(int? progresso = null)
    {
        return Task.CompletedTask;
    }
}

// WindowWrapper
// Adaptador necessário para passar o HWND do Word (Win32)
// como IWin32Window para MessageBox e Forms.
// Sem isso, MessageBox pode abrir atrás ou minimizado.
internal class WindowWrapper : IWin32Window //WindowWrapper Pa
{
    public IntPtr Handle { get; }

    public WindowWrapper(IntPtr handle)
    {
        Handle = handle;
    }
}
