using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 本地更新弹窗（现代扁平样式，自绘）。
    ///
    /// 为什么要有它：更新是"退出前最后一步"的操作，必须让用户看得见、能干预：
    ///   ① 显示下载进度 + 速度
    ///   ② 下载完、**安装前**检查当前打开的 PPT 有没有未保存的改动，并提供"保存全部"
    ///   ③ 保存确认之后才一键"卸载旧版 + 装新版"
    ///
    /// 样式：无边框 + 圆角 + 阴影 + 自绘控件（系统默认控件太"原始"，与插件 UI 不搭）。
    /// </summary>
    public class UpdateDialog : Form
    {
        // ---- 配色（与插件 Web UI 的蓝色系保持一致）----
        static readonly Color C_Bg = Color.White;
        static readonly Color C_Border = Color.FromArgb(229, 231, 235);
        static readonly Color C_Title = Color.FromArgb(17, 24, 39);
        static readonly Color C_Text = Color.FromArgb(55, 65, 81);
        static readonly Color C_Muted = Color.FromArgb(107, 114, 128);
        static readonly Color C_Primary = Color.FromArgb(37, 99, 235);
        static readonly Color C_PrimaryHover = Color.FromArgb(29, 78, 216);
        static readonly Color C_SoftBg = Color.FromArgb(249, 250, 251);
        static readonly Color C_Track = Color.FromArgb(229, 231, 235);
        static readonly Color C_Warn = Color.FromArgb(180, 83, 9);

        const int W = 580, H = 470;

        readonly UpdateResult _info;
        float _s = 1f;                  // DPI 缩放（按 96dpi 设计）

        TitleBar _bar;
        FlatButton _close;
        Label _title, _sub;
        Label _notesLabel;
        RoundedPanel _notesBox;
        RichTextBox _notes;
        SlimProgress _progress;
        Label _stage, _speed;
        FlatButton _go, _later;

        // 保存阶段
        Label _saveLabel;
        RoundedPanel _listBox;
        ListBox _list;
        FlatButton _saveAll, _skipSave, _cancel;
        SeparatorLine _sep;

        System.Windows.Forms.Timer _quitTimer;
        int _quitTries;
        string _installer, _dir;

        long _lastBytes;
        DateTime _lastTick = DateTime.MinValue;
        double _lastSpeed;

        public UpdateDialog(UpdateResult info)
        {
            _info = info;
            BuildUi();
        }

        int S(int px) { return (int)Math.Round(px * _s); }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= 0x20000;   // CS_DROPSHADOW
                return cp;
            }
        }

        // ---------------- 构建 ----------------

        void BuildUi()
        {
            Text = "素材库更新";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = true;
            BackColor = C_Bg;
            DoubleBuffered = true;
            Font = new Font("Microsoft YaHei UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;   // 自己按 DPI 缩放
            ClientSize = new Size(W, H);
            Padding = new Padding(1);

            _bar = new TitleBar { Dock = DockStyle.Top, Height = 44 };
            _bar.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) Drag(); };
            Controls.Add(_bar);

            _title = new Label
            {
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 20f, FontStyle.Bold),
                ForeColor = C_Title,
                BackColor = Color.Transparent,
            };
            Controls.Add(_title);

            _sub = new Label
            {
                AutoSize = false,
                ForeColor = C_Muted,
                BackColor = Color.Transparent,
                Font = new Font("Microsoft YaHei UI", 9.5f),
            };
            Controls.Add(_sub);

            _close = new FlatButton { Text = "✕", Kind = FlatButton.BtnKind.Ghost, Font = new Font("Segoe UI", 11f), ForeColor = C_Muted };
            _close.Click += (s, e) => { if (_quitTimer == null) Close(); };
            Controls.Add(_close);
            _bar.BringToFront();

            _notesLabel = new Label
            {
                Text = "更新内容",
                AutoSize = false,
                ForeColor = C_Muted,
                BackColor = Color.Transparent,
                Font = new Font("Microsoft YaHei UI", 9f),
            };
            Controls.Add(_notesLabel);

            _notesBox = new RoundedPanel { Radius = 10, Bg = C_SoftBg, BorderColor = C_Border };
            _notes = new RichTextBox
            {
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                BackColor = C_SoftBg,
                ForeColor = C_Text,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Font = new Font("Microsoft YaHei UI", 9f),
                DetectUrls = false,
                TabStop = false,
            };
            _notesBox.Controls.Add(_notes);
            Controls.Add(_notesBox);

            _progress = new SlimProgress { Track = C_Track, Fill = C_Primary, Radius = 4 };
            Controls.Add(_progress);

            _stage = new Label
            {
                AutoSize = false,
                ForeColor = C_Text,
                BackColor = Color.Transparent,
                Font = new Font("Microsoft YaHei UI", 9.5f),
            };
            Controls.Add(_stage);

            _speed = new Label
            {
                AutoSize = false,
                ForeColor = C_Muted,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("Microsoft YaHei UI", 9f),
            };
            Controls.Add(_speed);

            _go = new FlatButton { Text = "立即更新", Kind = FlatButton.BtnKind.Primary };
            _go.Click += (s, e) => StartDownload();
            Controls.Add(_go);

            _later = new FlatButton { Text = "稍后", Kind = FlatButton.BtnKind.Secondary };
            _later.Click += (s, e) => { if (_quitTimer == null) Close(); };
            Controls.Add(_later);

            _sep = new SeparatorLine();
            Controls.Add(_sep);

            // ---- 保存阶段 ----
            _saveLabel = new Label { AutoSize = false, ForeColor = C_Warn, BackColor = Color.Transparent, Visible = false };
            Controls.Add(_saveLabel);

            _listBox = new RoundedPanel { Radius = 10, Bg = C_SoftBg, BorderColor = C_Border, Visible = false };
            _list = new ListBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = C_SoftBg,
                ForeColor = C_Text,
                IntegralHeight = false,
                Font = new Font("Microsoft YaHei UI", 9f),
            };
            _listBox.Controls.Add(_list);
            Controls.Add(_listBox);

            _saveAll = new FlatButton { Text = "保存全部并继续", Kind = FlatButton.BtnKind.Primary, Visible = false };
            _saveAll.Click += (s, e) => SaveAllAndContinue();
            Controls.Add(_saveAll);

            _skipSave = new FlatButton { Text = "不保存，继续更新", Kind = FlatButton.BtnKind.Secondary, Visible = false };
            _skipSave.Click += (s, e) => InstallNow();
            Controls.Add(_skipSave);

            _cancel = new FlatButton { Text = "取消", Kind = FlatButton.BtnKind.Secondary, Visible = false };
            _cancel.Click += (s, e) => Close();
            Controls.Add(_cancel);

            _title.Text = "发现新版本 " + _info.LatestVersion;
            _sub.Text = "当前版本 " + _info.CurrentVersion + "　·　安装需要关闭 PowerPoint";
            var notes = (_info.Notes ?? "").Trim();
            _notes.Text = string.IsNullOrWhiteSpace(notes) || notes.StartsWith("（")
                ? "（本次发布没有填写更新说明）"
                : notes.Replace("\n", "\r\n").Replace("\r\r\n", "\r\n");
            _progress.Value = 0;
            _stage.Text = "点「立即更新」开始下载新版本";
            _speed.Text = "";
            LayoutPhase1();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            _s = DeviceDpi / 96f;
            ClientSize = new Size(S(W), S(H));
            ApplyRoundedRegion(S(14));
            LayoutPhase1();
        }

        void ApplyRoundedRegion(int radius)
        {
            try
            {
                using (var path = Rounded(new Rectangle(0, 0, Width, Height), radius))
                    Region = new Region(path);
            }
            catch { }
        }

        static GraphicsPath Rounded(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            int d = radius * 2;
            if (d <= 0) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        void LayoutPhase1()
        {
            int pad = S(26);
            _title.SetBounds(pad, S(56), S(W) - pad * 2, S(34));
            _sub.SetBounds(pad, S(96), S(W) - pad * 2, S(20));
            // 关闭按钮要**在标题栏之上**（BringToFront 已在 BuildUi 里做过），
            // 否则会被 TitleBar 这块 Dock=Top 的面板盖住而看不见。
            _close.SetBounds(S(W) - S(52), S(14), S(30), S(30));
            _close.BringToFront();

            _notesLabel.SetBounds(pad, S(130), S(W) - pad * 2, S(18));
            _notesBox.SetBounds(pad, S(152), S(W) - pad * 2, S(122));
            _notes.SetBounds(S(12), S(10), _notesBox.Width - S(24), _notesBox.Height - S(20));

            _progress.SetBounds(pad, S(292), S(W) - pad * 2, S(7));
            _stage.SetBounds(pad, S(308), S(360), S(22));
            _speed.SetBounds(S(W) - pad - S(200), S(308), S(200), S(22));

            _go.SetBounds(S(W) - pad - S(124), S(H) - S(24) - S(40), S(124), S(40));
            _later.SetBounds(_go.Left - S(10) - S(92), _go.Top, S(92), S(40));
            _sep.SetBounds(pad, _go.Top - S(18), S(W) - pad * 2, 1);
        }

        void LayoutPhase2()
        {
            int pad = S(26);
            _saveLabel.SetBounds(pad, S(130), S(W) - pad * 2, S(20));
            _listBox.SetBounds(pad, S(156), S(W) - pad * 2, S(174));
            _list.SetBounds(S(8), S(8), _listBox.Width - S(16), _listBox.Height - S(16));

            int y = S(H) - S(24) - S(40);
            _cancel.SetBounds(S(W) - pad - S(92), y, S(92), S(40));
            _skipSave.SetBounds(_cancel.Left - S(10) - S(148), y, S(148), S(40));
            _saveAll.SetBounds(_skipSave.Left - S(10) - S(156), y, S(156), S(40));
            _sep.SetBounds(pad, y - S(18), S(W) - pad * 2, 1);
        }

        // ---------------- 下载 ----------------

        void StartDownload()
        {
            _go.Enabled = false;
            _later.Text = "取消";
            _close.Enabled = false;
            _progress.Value = 0;
            _stage.Text = "正在连接…";
            _speed.Text = "";

            var t = new Thread(() =>
            {
                UpdateApply res;
                try { res = UpdateService.ApplyUpdateEx((stage, got, total) => Report(stage, got, total)); }
                catch (Exception ex) { res = new UpdateApply { Ok = false, Message = ex.Message }; }
                try { BeginInvoke(new Action(() => OnDownloadFinished(res))); } catch { }
            });
            t.IsBackground = true;
            t.Start();
        }

        void Report(string stage, long got, long total)
        {
            try
            {
                BeginInvoke(new Action(() =>
                {
                    if (stage == "download")
                    {
                        if (total > 0) _progress.Value = (int)Math.Min(100, got * 100 / total);

                        var now = DateTime.Now;
                        if (_lastTick != DateTime.MinValue)
                        {
                            var dt = (now - _lastTick).TotalSeconds;
                            if (dt >= 0.8)
                            {
                                _lastSpeed = (got - _lastBytes) / dt;
                                _lastBytes = got;
                                _lastTick = now;
                            }
                        }
                        else { _lastBytes = got; _lastTick = now; }

                        _stage.Text = total > 0
                            ? string.Format("正在下载…  {0:0.0} / {1:0.0} MB", got / 1048576.0, total / 1048576.0)
                            : string.Format("正在下载…  {0:0.0} MB", got / 1048576.0);
                        _speed.Text = _lastSpeed > 0 ? string.Format("{0:0.00} MB/s", _lastSpeed / 1048576.0) : "";
                    }
                    else if (stage == "verify")
                    {
                        _progress.Value = 100;
                        _stage.Text = "正在校验文件完整性…";
                        _speed.Text = "";
                    }
                    else if (stage == "extract")
                    {
                        _stage.Text = "正在解压…";
                        _speed.Text = "";
                    }
                }));
            }
            catch { }
        }

        void OnDownloadFinished(UpdateApply res)
        {
            _close.Enabled = true;
            if (!res.Ok)
            {
                _go.Enabled = true;
                _later.Text = "关闭";
                _stage.Text = "更新失败";
                _progress.Value = 0;
                MessageBox.Show(this,
                    res.Message + "\r\n\r\n也可到发布页手动下载。",
                    "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _installer = res.Installer;
            _dir = res.Dir;
            _progress.Value = 100;
            _stage.Text = "下载完成，准备安装…";
            CheckUnsavedAndAsk();
        }

        // ---------------- 未保存文稿 ----------------

        static bool IsDirty(Microsoft.Office.Interop.PowerPoint.Presentation p)
        {
            try { return p.Saved != Microsoft.Office.Core.MsoTriState.msoTrue; }
            catch { return false; }
        }

        void CheckUnsavedAndAsk()
        {
            var noPath = new List<string>();
            var dirty = new List<object>();
            try
            {
                var app = Globals.ThisAddIn != null ? Globals.ThisAddIn.Application : null;
                if (app != null)
                {
                    foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in app.Presentations)
                    {
                        try
                        {
                            if (!IsDirty(p)) continue;
                            if (string.IsNullOrEmpty(p.Path)) noPath.Add(p.Name);
                            else dirty.Add(p);
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex) { Log.Write("[update] 检查未保存文稿失败：" + ex.Message); }

            if (noPath.Count > 0)
            {
                MessageBox.Show(this,
                    "有 " + noPath.Count + " 个文稿**从未保存过**：\r\n\r\n  " + string.Join("\r\n  ", noPath.ToArray()) +
                    "\r\n\r\n请先在 PowerPoint 里手动保存它们（另存为），再重新点「立即更新」。\r\n" +
                    "（为了保证你的内容万无一失，这里不代为保存这类文稿。）",
                    "请先手动保存", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _stage.Text = "请先手动保存未命名文稿";
                _go.Enabled = true;
                _later.Text = "关闭";
                return;
            }

            if (dirty.Count == 0)
            {
                var r = MessageBox.Show(this,
                    "新版本 " + _info.LatestVersion + " 已下载完成。\r\n\r\n" +
                    "安装需要关闭 PowerPoint（新版文件要替换正在使用的旧版）。\r\n" +
                    "关闭后安装器会在后台完成升级，完成后重新打开 PowerPoint 即可。\r\n\r\n" +
                    "现在开始安装？",
                    "开始安装", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (r == DialogResult.OK) InstallNow();
                else { _go.Enabled = true; _later.Text = "稍后"; }
                return;
            }

            // 有未保存改动（已有路径）→ 列出文稿 + 「保存全部并继续」
            _notesLabel.Visible = _notesBox.Visible = _progress.Visible = _stage.Visible = _speed.Visible = false;
            _go.Visible = _later.Visible = false;

            _saveLabel.Text = "以下 " + dirty.Count + " 个文稿有未保存的改动，建议先保存（更新会关闭 PowerPoint）：";
            _saveLabel.Visible = true;
            _list.Items.Clear();
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in dirty)
            {
                try { _list.Items.Add("  " + p.Name + "    " + p.Path); } catch { }
            }
            _listBox.Visible = true;
            _saveAll.Visible = _skipSave.Visible = _cancel.Visible = true;
            _listBox.Tag = dirty;
            LayoutPhase2();
        }

        void SaveAllAndContinue()
        {
            var dirty = _listBox.Tag as List<object>;
            if (dirty == null) { InstallNow(); return; }

            _saveAll.Enabled = false;
            var failed = new List<string>();
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in dirty)
            {
                try { p.Save(); }
                catch (Exception ex) { try { failed.Add(p.Name + "：" + ex.Message); } catch { failed.Add(ex.Message); } }
            }
            _saveAll.Enabled = true;

            var still = new List<string>();
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in dirty)
            {
                try { if (IsDirty(p)) still.Add(p.Name); } catch { }
            }
            if (still.Count > 0)
            {
                MessageBox.Show(this,
                    "以下文稿仍未保存成功：\r\n\r\n  " + string.Join("\r\n  ", still.ToArray()) +
                    (failed.Count > 0 ? "\r\n\r\n错误信息：\r\n  " + string.Join("\r\n  ", failed.ToArray()) : "") +
                    "\r\n\r\n请手动保存后再点「保存全部并继续」，或选择「不保存，继续更新」。",
                    "仍有未保存的文稿", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Log.Write("[update] 已保存 " + dirty.Count + " 个文稿，继续安装");
            InstallNow();
        }

        // ---------------- 安装 ----------------

        void InstallNow()
        {
            if (string.IsNullOrEmpty(_installer) || !File.Exists(_installer))
            {
                MessageBox.Show(this, "安装器不存在，无法自动更新。", "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _saveLabel.Visible = _listBox.Visible = false;
            _saveAll.Visible = _skipSave.Visible = _cancel.Visible = false;
            _notesLabel.Visible = _notesBox.Visible = _progress.Visible = _stage.Visible = _speed.Visible = true;
            _progress.Indeterminate = true;
            _later.Visible = false;
            _go.Visible = false;
            _close.Visible = true;
            _close.Enabled = false;

            _title.Text = "正在安装 " + _info.LatestVersion;
            _sub.Text = "安装器已启动，PowerPoint 将自动关闭并在后台完成升级";
            _stage.Text = "正在启动安装器…";
            _speed.Text = "";

            try
            {
                var psi = new ProcessStartInfo(_installer, "--pending-update \"" + _dir + "\"");
                psi.UseShellExecute = true;
                psi.WorkingDirectory = Path.GetDirectoryName(_installer);
                Process.Start(psi);
                Log.Write("[update] 已启动安装器：" + _installer + " --pending-update " + _dir);
            }
            catch (Exception ex)
            {
                Log.Write("[update] 启动安装器失败：" + ex.Message);
                MessageBox.Show(this, "启动安装器失败：" + ex.Message, "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _quitTries = 0;
            _quitTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            _quitTimer.Tick += (s, e) =>
            {
                _quitTries++;
                try
                {
                    var app = Globals.ThisAddIn != null ? Globals.ThisAddIn.Application : null;
                    if (app != null) app.Quit();
                }
                catch (Exception ex) { Log.Write("[update] 关闭 PowerPoint 失败：" + ex.Message); }

                if (_quitTries >= 6)
                {
                    _quitTimer.Stop();
                    _quitTimer.Dispose();
                    _quitTimer = null;
                    _progress.Indeterminate = false;
                    _progress.Value = 100;
                    _close.Enabled = true;

                    // 关键判断：**有没有"有窗口"的 PowerPoint**。
                    // 没有窗口（僵尸进程 / 宿主本来就没界面）时，安装器的同一判据也不会等它，
                    // 所以**不能再催用户去关 PowerPoint** —— 用户会找不到可关的东西，
                    // 以为卡死了（这正是用户反馈的"没有 PPT 却让我关 PPT"）。
                    if (!UpdateService.HasWindowedPowerPoint())
                    {
                        _stage.Text = "PowerPoint 已关闭，安装器正在后台完成升级…";
                        _later.Visible = true;
                        _later.Text = "关闭本窗口";
                        _later.SetBounds(_go.Left, _go.Top, S(120), S(38));
                        Log.Write("[update] 无可见 PowerPoint 窗口 → 安装器会直接安装（无需用户操作）");

                        // 5 秒后自动关掉本窗口：更新已经交给安装器，别让用户以为要一直等
                        var auto = new System.Windows.Forms.Timer { Interval = 5000 };
                        auto.Tick += (s2, e2) =>
                        {
                            auto.Stop();
                            auto.Dispose();
                            try { Close(); } catch { }
                        };
                        auto.Start();
                    }
                    else
                    {
                        _stage.Text = "请手动关闭 PowerPoint —— 安装器正在等待，关闭后会自动完成升级";
                        _later.Visible = true;
                        _later.Text = "关闭本窗口";
                        _later.SetBounds(_go.Left, _go.Top, S(120), S(38));
                        Log.Write("[update] 自动关闭未生效 → 请手动关闭 PowerPoint，安装器会自动完成升级");
                    }
                }
            };
            _quitTimer.Start();
        }

        // ---------------- 拖动无边框窗口 ----------------

        [DllImport("user32.dll")] static extern bool ReleaseCapture();
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr w, IntPtr l);

        void Drag()
        {
            ReleaseCapture();
            SendMessage(Handle, 0xA1, (IntPtr)0x2, IntPtr.Zero);   // WM_NCLBUTTONDOWN / HTCAPTION
        }

        // ---------------- 自绘控件 ----------------

        /// <summary>标题栏（透明，只为接收拖动与承载关闭按钮）。</summary>
        class TitleBar : Panel
        {
            public TitleBar()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
                BackColor = C_Bg;
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(C_Border))
                    e.Graphics.DrawLine(pen, 1, Height - 1, Width - 2, Height - 1);
            }
        }

        /// <summary>扁平按钮（圆角、悬停/按下/禁用态）。</summary>
        class FlatButton : Control
        {
            public enum BtnKind { Primary, Secondary, Ghost }
            BtnKind _kind = BtnKind.Secondary;
            bool _hover, _down;

            public FlatButton()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                         | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                Cursor = Cursors.Hand;
                Font = new Font("Microsoft YaHei UI", 9.5f);
                BackColor = C_Bg;
            }

            public BtnKind Kind { get { return _kind; } set { _kind = value; Invalidate(); } }

            protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnMouseDown(MouseEventArgs e) { _down = true; Invalidate(); base.OnMouseDown(e); }
            protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e); }
            protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new Rectangle(0, 0, Width - 1, Height - 1);
                int radius = Math.Min(8, Height / 3);

                Color fill, border, text;
                switch (_kind)
                {
                    case BtnKind.Primary:
                        fill = !Enabled ? Color.FromArgb(147, 197, 253)
                             : _down ? Color.FromArgb(30, 64, 175)
                             : _hover ? C_PrimaryHover : C_Primary;
                        border = fill; text = Color.White; break;
                    case BtnKind.Ghost:
                        fill = _down ? Color.FromArgb(243, 244, 246) : (_hover ? C_SoftBg : C_Bg);
                        border = Color.Transparent; text = C_Muted; break;
                    default:
                        fill = _down ? Color.FromArgb(243, 244, 246) : (_hover ? C_SoftBg : C_Bg);
                        border = C_Border; text = C_Text; break;
                }

                using (var path = Rounded(r, radius))
                {
                    using (var b = new SolidBrush(fill)) g.FillPath(b, path);
                    if (border != Color.Transparent)
                        using (var p = new Pen(border)) g.DrawPath(p, path);
                }
                TextRenderer.DrawText(g, Text, Font, r,
                    Enabled ? text : Color.FromArgb(156, 163, 175),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
        }

        /// <summary>1px 分隔线。</summary>
        class SeparatorLine : Control
        {
            public SeparatorLine()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                         | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                BackColor = C_Bg;
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                using (var p = new Pen(C_Border))
                    e.Graphics.DrawLine(p, 0, 0, Width, 0);
            }
        }

        /// <summary>细长进度条（圆角，可切换不确定态）。</summary>
        class SlimProgress : Control
        {
            int _value;
            public Color Track = C_Track;
            public Color Fill = C_Primary;
            public int Radius = 4;
            public bool Indeterminate;

            int _march;
            readonly System.Windows.Forms.Timer _anim;

            public SlimProgress()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                         | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                BackColor = C_Bg;
                _anim = new System.Windows.Forms.Timer { Interval = 60 };
                _anim.Tick += (s, e) => { _march = (_march + 6) % Math.Max(1, Width); Invalidate(); };
            }

            public int Value
            {
                get { return _value; }
                set { _value = Math.Max(0, Math.Min(100, value)); Invalidate(); }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                if (Indeterminate && !_anim.Enabled) _anim.Start();
                if (!Indeterminate && _anim.Enabled) _anim.Stop();

                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                int h = Height, radius = Math.Min(Radius, h / 2);
                var full = new Rectangle(0, 0, Math.Max(1, Width - 1), Math.Max(1, h - 1));

                using (var path = Rounded(full, radius))
                using (var b = new SolidBrush(Track))
                    g.FillPath(b, path);

                if (Indeterminate)
                {
                    int w = Math.Max(40, Width / 4);
                    int x = _march - w;
                    var clip = new Rectangle(Math.Max(0, x), 0, w, h);
                    var r2 = Rectangle.Intersect(full, clip);
                    if (r2.Width > 2)
                    {
                        var st = g.Save();
                        using (var path = Rounded(full, radius)) g.SetClip(path);
                        using (var b = new SolidBrush(Fill)) g.FillRectangle(b, r2);
                        g.Restore(st);
                    }
                    return;
                }

                int fw = (int)Math.Round((Width - 1) * (_value / 100.0));
                if (fw > 1)
                {
                    var fr = new Rectangle(0, 0, fw, Math.Max(1, h - 1));
                    using (var path = Rounded(fr, radius))
                    using (var b = new SolidBrush(Fill))
                        g.FillPath(b, path);
                }
            }
        }

        /// <summary>圆角面板（浅底 + 细边框），承载说明文本/文稿列表。</summary>
        class RoundedPanel : Panel
        {
            public int Radius = 10;
            public Color Bg = C_SoftBg;
            public Color BorderColor = C_Border;
            public RoundedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint
                         | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                BackColor = C_Bg;
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Rounded(new Rectangle(0, 0, Width - 1, Height - 1), Radius))
                {
                    using (var b = new SolidBrush(Bg)) g.FillPath(b, path);
                    using (var p = new Pen(BorderColor)) g.DrawPath(p, path);
                }
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = Rounded(new Rectangle(0, 0, Width - 1, Height - 1), S(14)))
            using (var p = new Pen(C_Border))
                e.Graphics.DrawPath(p, path);
        }

        // ---------------- 入口 ----------------

        static UpdateDialog _open;

        /// <summary>
        /// 显示更新弹窗（供桥接调用）。
        ///
        /// ⚠️ 两个关键点（都踩过）：
        ///   · **不能在 WebView2 的消息回调里直接 Show()/ShowDialog()** —— 回调运行在嵌套消息循环里，
        ///     实测窗体构造成功但窗口不显示、调用也永不返回（前端按钮卡在"打开中…"）。
        ///     必须把显示动作**投递到 UI 消息队列**（PostToUi），等当前回调返回后再显示。
        ///   · 用非模态 Show()，不阻塞 PowerPoint 的消息处理。
        /// </summary>
        public static string ShowForUpdate()
        {
            Log.Write("[update] ShowForUpdate begin");
            var info = UpdateService.Check();
            if (info == null || !info.Ok)
                return "检查更新失败：" + (info != null ? info.Error : "未知");
            if (!info.HasUpdate) return "当前已是最新版本 " + info.CurrentVersion + "。";
            ShowWith(info);
            return null;
        }

        /// <summary>直接用已有结果打开弹窗（避免在 UI 线程上再次联网检查）。</summary>
        public static void ShowWith(UpdateResult info)
        {
            Log.Write("[update] ShowWith cur=" + info.CurrentVersion + " latest=" + info.LatestVersion);

            if (_open != null && !_open.IsDisposed)
            {
                try { _open.Activate(); } catch { }
                return;
            }

            BridgeController.PostToUiStatic(() =>
            {
                try
                {
                    if (_open != null && !_open.IsDisposed) { _open.Activate(); return; }
                    var dlg = new UpdateDialog(info);
                    dlg.FormClosed += (s, e) => { _open = null; };
                    _open = dlg;
                    dlg.Show();
                    Log.Write("[update] 更新弹窗已显示");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "UpdateDialog.Show");
                    MessageBox.Show("打开更新窗口失败：" + ex.Message, "素材库更新",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
        }
    }
}
