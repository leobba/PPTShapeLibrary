using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 本地更新弹窗（C# WinForms，不依赖 WebView 面板）。
    ///
    /// 为什么要有它：更新是"退出前最后一步"的操作，必须让用户看得见、能干预：
    ///   ① 显示下载进度 + 速度（不然点完只看到一个"准备中…"，用户不知道在干什么）
    ///   ② 下载完、**安装前**检查当前打开的 PPT 有没有未保存的改动，并提供"保存全部"按钮
    ///      （否则 PowerPoint 一关，用户的改动就没了 —— 这是最不能接受的事故）
    ///   ③ 保存确认之后才一键"卸载旧版 + 装新版"
    ///
    /// 流程：检查版本 → 下载(进度/速度) → SHA256 校验 → 解压 → 【检查未保存文稿】→ 启动安装器 → 关闭 PowerPoint
    /// </summary>
    public class UpdateDialog : Form
    {
        private readonly UpdateResult _info;

        private Label _title;
        private TextBox _notes;
        private ProgressBar _bar;
        private Label _stage;
        private Label _speed;
        private Button _go;
        private Button _close;

        private Panel _savePanel;
        private Label _saveHint;
        private ListBox _saveList;
        private Button _saveAll;
        private Button _skipSave;
        private Button _cancelAll;

        private System.Windows.Forms.Timer _quitTimer;
        private int _quitTries;
        private string _installer, _dir;

        private long _lastBytes;
        private DateTime _lastTick = DateTime.MinValue;
        private double _lastSpeed;      // 字节/秒

        public UpdateDialog(UpdateResult info)
        {
            _info = info;
            Log.Write("[update] UpdateDialog 构造开始");
            BuildUi();
            Log.Write("[update] UpdateDialog 构造完成");
        }

        // ---------------- UI ----------------

        private void BuildUi()
        {
            Text = "素材库更新";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 420);
            Font = new Font("Microsoft YaHei UI", 9F);
            TopMost = true;

            _title = new Label
            {
                Text = "发现新版本 " + _info.LatestVersion + "（当前 " + _info.CurrentVersion + "）",
                Location = new Point(16, 14),
                Size = new Size(528, 24),
                Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold),
            };
            Controls.Add(_title);

            Controls.Add(new Label
            {
                Text = "更新内容：",
                Location = new Point(16, 44),
                Size = new Size(200, 18),
                ForeColor = Color.FromArgb(90, 100, 110),
            });

            _notes = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new Point(16, 64),
                Size = new Size(528, 190),
                BackColor = Color.FromArgb(248, 249, 250),
                Text = (string.IsNullOrWhiteSpace(_info.Notes) ? "（本次发布没有填写更新说明）" : _info.Notes),
            };
            Controls.Add(_notes);

            _bar = new ProgressBar
            {
                Location = new Point(16, 266),
                Size = new Size(528, 20),
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
            };
            Controls.Add(_bar);

            _stage = new Label
            {
                Text = "等待开始…",
                Location = new Point(16, 292),
                Size = new Size(300, 20),
            };
            Controls.Add(_stage);

            _speed = new Label
            {
                Text = "",
                Location = new Point(316, 292),
                Size = new Size(228, 20),
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.FromArgb(90, 100, 110),
            };
            Controls.Add(_speed);

            // 未保存文稿面板（默认隐藏）
            _savePanel = new Panel { Location = new Point(16, 256), Size = new Size(528, 120), Visible = false };
            _saveHint = new Label
            {
                Text = "以下文稿有未保存的改动，建议先保存再更新：",
                Location = new Point(0, 0),
                Size = new Size(528, 20),
                ForeColor = Color.FromArgb(190, 90, 0),
            };
            _savePanel.Controls.Add(_saveHint);
            _saveList = new ListBox { Location = new Point(0, 24), Size = new Size(528, 92) };
            _savePanel.Controls.Add(_saveList);
            Controls.Add(_savePanel);

            _go = new Button
            {
                Text = "立即更新",
                Location = new Point(336, 366),
                Size = new Size(120, 34),
            };
            _go.Click += (s, e) => StartDownload();
            Controls.Add(_go);

            _close = new Button
            {
                Text = "稍后",
                Location = new Point(462, 366),
                Size = new Size(82, 34),
            };
            _close.Click += (s, e) =>
            {
                if (_quitTimer != null) return;   // 安装中不允许关
                Close();
            };
            Controls.Add(_close);

            // 保存阶段的按钮（默认隐藏）
            _saveAll = new Button { Text = "保存全部并继续", Location = new Point(300, 366), Size = new Size(140, 34), Visible = false };
            _saveAll.Click += (s, e) => SaveAllAndContinue();
            Controls.Add(_saveAll);

            _skipSave = new Button { Text = "不保存，继续更新", Location = new Point(170, 366), Size = new Size(124, 34), Visible = false };
            _skipSave.Click += (s, e) => InstallNow();
            Controls.Add(_skipSave);

            _cancelAll = new Button { Text = "取消", Location = new Point(462, 366), Size = new Size(82, 34), Visible = false };
            _cancelAll.Click += (s, e) => Close();
            Controls.Add(_cancelAll);
        }

        // ---------------- 下载 ----------------

        private void StartDownload()
        {
            _go.Enabled = false;
            _close.Text = "取消";
            _bar.Value = 0;
            _stage.Text = "正在连接…";
            _speed.Text = "";

            var t = new Thread(() =>
            {
                UpdateApply res;
                try
                {
                    res = UpdateService.ApplyUpdateEx((stage, got, total) => Report(stage, got, total));
                }
                catch (Exception ex)
                {
                    res = new UpdateApply { Ok = false, Message = ex.Message };
                }
                try { BeginInvoke(new Action(() => OnDownloadFinished(res))); } catch { }
            });
            t.IsBackground = true;
            t.Start();
        }

        private void Report(string stage, long got, long total)
        {
            try
            {
                BeginInvoke(new Action(() =>
                {
                    if (stage == "download")
                    {
                        if (total > 0)
                        {
                            int pct = (int)Math.Min(100, got * 100 / total);
                            if (pct != _bar.Value) _bar.Value = pct;
                        }
                        // 速度：按 ~1 秒窗口算瞬时速度，避免数字乱跳
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
                            ? string.Format("正在下载… {0:0.0} / {1:0.0} MB", got / 1048576.0, total / 1048576.0)
                            : string.Format("正在下载… {0:0.0} MB", got / 1048576.0);
                        _speed.Text = _lastSpeed > 0
                            ? string.Format("{0:0.00} MB/s", _lastSpeed / 1048576.0)
                            : "";
                    }
                    else if (stage == "verify")
                    {
                        _bar.Value = 100;
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

        private void OnDownloadFinished(UpdateApply res)
        {
            if (!res.Ok)
            {
                _go.Enabled = true;
                _close.Text = "关闭";
                _stage.Text = "更新失败";
                _bar.Value = 0;
                MessageBox.Show(this,
                    res.Message + "\r\n\r\n也可点面板里的「查看更新」到发布页手动下载。",
                    "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _installer = res.Installer;
            _dir = res.Dir;
            _bar.Value = 100;
            _stage.Text = "下载完成，准备安装…";
            CheckUnsavedAndAsk();
        }

        // ---------------- 未保存文稿检查 ----------------

        /// <summary>文稿是否有未保存改动。注意 PowerPoint 的 Presentation.Saved 是 MsoTriState 而非 bool。</summary>
        private static bool IsDirty(Microsoft.Office.Interop.PowerPoint.Presentation p)
        {
            try { return p.Saved != Microsoft.Office.Core.MsoTriState.msoTrue; }
            catch { return false; }   // 读不到就当作"没有改动"，避免误拦
        }

        private void CheckUnsavedAndAsk()
        {
            var dirtyNoPath = new List<string>();
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
                            var name = p.Name;
                            if (string.IsNullOrEmpty(p.Path))
                            {
                                dirtyNoPath.Add(name);
                            }
                            else
                            {
                                dirty.Add(p);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write("[update] 检查未保存文稿失败：" + ex.Message);
            }

            // 从未保存过的新文稿：自动保存会弹"另存为"（而且会被本弹窗挡在后面），只能请用户手动处理
            if (dirtyNoPath.Count > 0)
            {
                MessageBox.Show(this,
                    "有 " + dirtyNoPath.Count + " 个文稿**从未保存过**：\r\n\r\n  " + string.Join("\r\n  ", dirtyNoPath.ToArray()) +
                    "\r\n\r\n请先在 PowerPoint 里手动保存它们（另存为），再重新点「立即更新」。\r\n" +
                    "（为了保证你的内容万无一失，这里不代为保存这类文稿。）",
                    "请先手动保存", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _stage.Text = "请先手动保存未命名文稿";
                _go.Enabled = true;
                _close.Text = "关闭";
                return;
            }

            if (dirty.Count == 0)
            {
                // 没有未保存改动 → 直接确认安装
                var r = MessageBox.Show(this,
                    "新版本 " + _info.LatestVersion + " 已下载完成。\r\n\r\n" +
                    "安装需要关闭 PowerPoint（新版文件要替换正在使用的旧版）。\r\n" +
                    "关闭后安装器会在后台完成升级，完成后重新打开 PowerPoint 即可。\r\n\r\n" +
                    "现在开始安装？",
                    "开始安装", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (r == DialogResult.OK) InstallNow();
                else { _go.Enabled = true; _close.Text = "稍后"; }
                return;
            }

            // 有未保存改动（已有路径）→ 列出来，给"保存全部"按钮
            _bar.Visible = false;
            _stage.Visible = false;
            _speed.Visible = false;
            _go.Visible = false;
            _close.Visible = false;

            _saveHint.Text = "以下 " + dirty.Count + " 个文稿有**未保存**的改动，建议先保存再更新（更新会关闭 PowerPoint）：";
            _saveList.Items.Clear();
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in dirty)
            {
                try { _saveList.Items.Add(p.Name + "    " + p.Path); } catch { }
            }
            _savePanel.Visible = true;
            _saveAll.Visible = true;
            _skipSave.Visible = true;
            _cancelAll.Visible = true;
            _savePanel.Tag = dirty;   // 存起来给"保存全部"用
        }

        private void SaveAllAndContinue()
        {
            var dirty = _savePanel.Tag as List<object>;
            if (dirty == null) { InstallNow(); return; }

            _saveAll.Enabled = false;
            var failed = new List<string>();
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in dirty)
            {
                try { p.Save(); }
                catch (Exception ex)
                {
                    try { failed.Add(p.Name + "：" + ex.Message); } catch { failed.Add(ex.Message); }
                }
            }
            _saveAll.Enabled = true;

            // 复查：仍有未保存的（例如被取消/只读）就停下来告诉用户
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

        private void InstallNow()
        {
            if (string.IsNullOrEmpty(_installer) || !File.Exists(_installer))
            {
                MessageBox.Show(this, "安装器不存在，无法自动更新。", "更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _saveAll.Visible = false;
            _skipSave.Visible = false;
            _cancelAll.Visible = false;
            _savePanel.Visible = false;

            _bar.Visible = true;
            _stage.Visible = true;
            _speed.Visible = true;
            _bar.Style = ProgressBarStyle.Marquee;
            _stage.Text = "正在启动安装器，PowerPoint 即将关闭…";
            _speed.Text = "";
            _close.Visible = false;

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

            // 关闭 PowerPoint（重试，见 ScheduleQuit）
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
                catch (Exception ex)
                {
                    Log.Write("[update] 关闭 PowerPoint 失败：" + ex.Message);
                }
                if (_quitTries >= 6)
                {
                    _quitTimer.Stop();
                    _quitTimer.Dispose();
                    _quitTimer = null;
                    // 关不掉：告诉用户手动关，安装器会一直等
                    _bar.Style = ProgressBarStyle.Continuous;
                    _bar.Value = 100;
                    _stage.Text = "请手动关闭 PowerPoint —— 安装器正在等待，关闭后会自动完成升级";
                    _close.Visible = true;
                    _close.Text = "关闭本窗口";
                    Log.Write("[update] 自动关闭未生效 → 请手动关闭 PowerPoint，安装器会自动完成升级");
                    MessageBox.Show(this,
                        "PowerPoint 没有自动关闭。\r\n\r\n" +
                        "请**手动关闭 PowerPoint**，安装器会一直等待（最多 3 分钟），关掉后自动完成升级。",
                        "请手动关闭 PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            _quitTimer.Start();
        }

        private static UpdateDialog _open;   // 保引用，避免被 GC 回收；同时也用于单实例

        /// <summary>
        /// 显示更新弹窗（供桥接调用）。
        ///
        /// ⚠️ 两个关键点（都踩过）：
        ///   · **不能在 WebView2 的消息回调里直接 ShowDialog()** —— 回调运行在嵌套消息循环里，
        ///     实测窗体构造成功但 ShowDialog 之后窗口不显示、调用也永不返回（前端按钮卡在"打开中…"）。
        ///     正确做法是把显示动作**投递到 UI 消息队列**（PostToUi），等当前回调返回后再显示。
        ///   · 用**非模态** Show()，不阻塞 PowerPoint 的消息处理（下载期间界面仍然可响应）。
        ///
        /// 返回 null 表示已安排显示；否则是"不能更新"的原因（已是最新 / 检查失败）。
        /// </summary>
        public static string ShowForUpdate()
        {
            Log.Write("[update] ShowForUpdate begin");
            var info = UpdateService.Check();
            Log.Write("[update] 检查完成 Ok=" + info.Ok + " HasUpdate=" + info.HasUpdate +
                      " cur=" + info.CurrentVersion + " latest=" + info.LatestVersion);
            if (!info.Ok) return "检查更新失败：" + info.Error;
            if (!info.HasUpdate) return "当前已是最新版本 " + info.CurrentVersion + "。";

            // 已开着就把它激活，不重复开
            if (_open != null && !_open.IsDisposed)
            {
                try { _open.Activate(); } catch { }
                return null;
            }

            // 投递到 UI 队列后再显示（不要在 WebView2 回调栈里 Show）
            BridgeController.PostToUiStatic(() =>
            {
                try
                {
                    if (_open != null && !_open.IsDisposed) { _open.Activate(); return; }
                    var dlg = new UpdateDialog(info);
                    dlg.FormClosed += (s, e) => { _open = null; };
                    _open = dlg;
                    dlg.Show();          // 非模态
                    Log.Write("[update] 更新弹窗已显示");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "UpdateDialog.Show");
                    MessageBox.Show("打开更新窗口失败：" + ex.Message, "素材库更新",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
            return null;
        }
    }
}
