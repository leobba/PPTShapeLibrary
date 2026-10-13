using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 长任务（整稿收藏 / 重建预览 / 整稿导出…）的统一上下文：
    ///
    /// 背景：桥消息在 PowerPoint 的主 STA 线程上同步执行，长循环期间该线程被占满，
    /// 前端既看不到进度、也送不进"取消"（消息排在队列里没人处理）。
    /// 这里提供三件事：
    ///   1. <see cref="Report"/> —— 通过桥把进度推给前端（PostWebMessageAsJson 直接投递到渲染进程，
    ///      即使宿主 UI 线程被占，前端仍能收到并刷新进度条）；
    ///   2. <see cref="Pump"/> —— 每 ~150ms 调一次 Application.DoEvents()，让队列里的
    ///      "cancel" 消息有机会被处理（否则取消永远无效）；配套 <see cref="Cancelled"/>；
    ///   3. <see cref="RegisterComFilter"/> / <see cref="RevokeComFilter"/> —— 在长任务期间注册
    ///      IOleMessageFilter。DoEvents 会让用户有机会在操作中途点击 PowerPoint，
    ///      此后我们的 COM 调用可能被拒（RPC_E_CALL_REJECTED / 0x80010001），
    ///      消息过滤器负责自动重试而不是把异常抛给用户。
    ///
    /// 注意：DoEvents 会引入重入风险，因此 BridgeController 在长任务期间会拒绝除
    /// cancel/ping 之外的任何新请求（见 <see cref="BridgeController"/> 的 busy 判断）。
    /// </summary>
    public class LongOperation
    {
        private readonly Action<string, int, int> _report;   // (reqId, current, total) 由桥注入

        public string ReqId { get; private set; }
        public string Title { get; private set; }
        public volatile bool Cancelled = false;

        private bool _filterRegistered;

        /// <summary>取消通道的 URL（形如 http://127.0.0.1:52341/cancel?t=xxxx），
        /// 随 progress 一起下发；前端点"取消"时 fetch 它。
        ///
        /// 为什么不用普通桥消息：长任务期间宿主 UI 线程被占，实测
        /// **Application.DoEvents() 也无法让 WebView2 的 WebMessageReceived 送达**
        /// （实测：循环里 DoEvents 调用了 6~7 次，期间收到的桥消息数为 0），
        /// 所以取消必须走一条不依赖 UI 消息循环的通道。既然 DoEvents 无用，
        /// 也就不再泵消息——避免"长任务中途被用户操作重入"的风险。</summary>
        public string CancelUrl { get; private set; }

        private CancelChannel _channel;

        public LongOperation(string reqId, string title, Action<string, int, int> report)
        {
            ReqId = reqId ?? "";
            Title = title ?? "";
            _report = report;
        }

        /// <summary>开启取消通道（仅监听 127.0.0.1 的临时端口，任务结束即关闭）。</summary>
        public void OpenCancelChannel()
        {
            try
            {
                _channel = new CancelChannel(() => { Cancelled = true; });
                CancelUrl = _channel.Start();
            }
            catch (Exception ex)
            {
                Log.Write("cancel channel unavailable: " + ex.Message);
                CancelUrl = null;
            }
        }

        public void CloseCancelChannel()
        {
            try { if (_channel != null) _channel.Stop(); } catch { }
            _channel = null;
            CancelUrl = null;
        }

        /// <summary>上报进度（total&lt;=0 表示不确定进度）。</summary>
        public void Report(int current, int total)
        {
            if (_report == null) return;
            try { _report(ReqId, current, total); } catch { }
        }

        /// <summary>上报进度并按需返回取消状态。调用方在长循环里逐项调用。</summary>
        public bool Step(int current, int total)
        {
            Report(current, total);
            return Cancelled;
        }

        /// <summary>长任务开始时注册消息过滤器（失败静默：没有它也能跑，只是遇到拒绝调用会直接报错）。</summary>
        public void RegisterComFilter()
        {
            if (_filterRegistered) return;
            try { ComMessageFilter.Register(); _filterRegistered = true; }
            catch (Exception ex) { Log.Write("MessageFilter register failed: " + ex.Message); }
        }

        /// <summary>长任务结束时务必调用，恢复 PowerPoint 自己的消息过滤器。</summary>
        public void RevokeComFilter()
        {
            if (!_filterRegistered) return;
            _filterRegistered = false;
            try { ComMessageFilter.Revoke(); }
            catch (Exception ex) { Log.Write("MessageFilter revoke failed: " + ex.Message); }
        }
    }

    /// <summary>
    /// 极简取消通道：只监听 127.0.0.1 的一个临时端口，只认一条
    /// `GET /cancel?t=&lt;随机令牌&gt;`，收到即置取消标志。任务开始才开、结束就关。
    ///
    /// 为什么不用 HTTP.sys(HttpListener)：绑定 URL 前缀通常需要管理员或 netsh urlacl；
    /// 裸 TcpListener 只是普通套接字，普通用户即可绑定回环端口。
    /// 安全边界：仅回环地址 + 随机端口 + 随机令牌 + 只认 cancel + 只在长任务期间存活。
    /// </summary>
    internal sealed class CancelChannel
    {
        private readonly Action _onCancel;
        private readonly string _token = Guid.NewGuid().ToString("N");
        private System.Net.Sockets.TcpListener _listener;
        private volatile bool _running;
        private System.Threading.Thread _thread;

        public CancelChannel(Action onCancel)
        {
            _onCancel = onCancel;
        }

        public string Start()
        {
            _listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
            _listener.Start();
            int port = ((System.Net.IPEndPoint)_listener.LocalEndpoint).Port;
            _running = true;
            _thread = new System.Threading.Thread(AcceptLoop) { IsBackground = true, Name = "PPTSL-CancelChannel" };
            _thread.Start();
            return "http://127.0.0.1:" + port + "/cancel?t=" + _token;
        }

        public void Stop()
        {
            _running = false;
            try { if (_listener != null) _listener.Stop(); } catch { }
            _listener = null;
        }

        private void AcceptLoop()
        {
            while (_running)
            {
                System.Net.Sockets.TcpClient client = null;
                try
                {
                    client = _listener.AcceptTcpClient();
                }
                catch
                {
                    return;   // Stop() 会让 Accept 抛异常，正常退出
                }
                try
                {
                    client.ReceiveTimeout = 2000;
                    client.SendTimeout = 2000;
                    using (var stream = client.GetStream())
                    {
                        var buf = new byte[2048];
                        int n = stream.Read(buf, 0, buf.Length);
                        string req = n > 0 ? System.Text.Encoding.ASCII.GetString(buf, 0, n) : "";
                        bool ok = req.StartsWith("GET /cancel", StringComparison.OrdinalIgnoreCase)
                                  && req.IndexOf("t=" + _token, StringComparison.Ordinal) >= 0;
                        if (ok)
                        {
                            Log.Write("[bridge] cancel via channel");
                            try { _onCancel(); } catch { }
                        }
                        string body = ok ? "cancelling" : "bad request";
                        string resp = "HTTP/1.1 " + (ok ? "200 OK" : "403 Forbidden") + "\r\n" +
                                      "Access-Control-Allow-Origin: *\r\n" +
                                      "Access-Control-Allow-Methods: GET, OPTIONS\r\n" +
                                      "Content-Type: text/plain; charset=utf-8\r\n" +
                                      "Content-Length: " + System.Text.Encoding.UTF8.GetByteCount(body) + "\r\n" +
                                      "Connection: close\r\n\r\n" + body;
                        var outBuf = System.Text.Encoding.UTF8.GetBytes(resp);
                        stream.Write(outBuf, 0, outBuf.Length);
                    }
                }
                catch { }
                finally
                {
                    try { client.Close(); } catch { }
                }
            }
        }
    }

    /// <summary>
    /// IOleMessageFilter：Office 忙时（用户中途点了 PowerPoint、弹出对话框等）自动重试被拒的 COM 调用，
    /// 而不是抛 "Call was rejected by callee"。
    /// 只在长任务期间注册、结束立即恢复，避免长期顶掉 PowerPoint 自己的过滤器。
    /// </summary>
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    internal static class ComMessageFilter
    {
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        private const int SERVERCALL_ISHANDLED = 0;
        private const int SERVERCALL_RETRYLATER = 2;
        private const int PENDINGMSG_WAITDEFPROCESS = 2;

        private static IOleMessageFilter _filter;      // 必须保持引用，否则 GC 后回调会崩
        private static IOleMessageFilter _previous;    // PowerPoint 原有的过滤器，用于恢复
        private static int _retryDeadline;

        private sealed class Filter : IOleMessageFilter
        {
            public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                return SERVERCALL_ISHANDLED;
            }

            public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == SERVERCALL_RETRYLATER)
                {
                    // 最多重试 5 秒；超时返回 -1 让调用方拿到真实异常（而不是无限卡住）
                    if (_retryDeadline == 0) _retryDeadline = dwTickCount + 5000;
                    if (dwTickCount > _retryDeadline) { _retryDeadline = 0; return -1; }
                    return 100;   // 100ms 后重试
                }
                _retryDeadline = 0;
                return -1;
            }

            public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            {
                return PENDINGMSG_WAITDEFPROCESS;
            }
        }

        public static void Register()
        {
            if (_filter != null) return;
            _filter = new Filter();
            _retryDeadline = 0;
            IOleMessageFilter old;
            int hr = CoRegisterMessageFilter(_filter, out old);
            if (hr != 0) { _filter = null; Marshal.ThrowExceptionForHR(hr); }
            _previous = old;
        }

        public static void Revoke()
        {
            if (_filter == null) return;
            try
            {
                IOleMessageFilter dummy;
                CoRegisterMessageFilter(_previous, out dummy);   // 恢复 PowerPoint 自己的过滤器
            }
            finally
            {
                _filter = null;
                _previous = null;
                _retryDeadline = 0;
            }
        }
    }
}
