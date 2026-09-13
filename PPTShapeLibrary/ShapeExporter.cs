using System;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 收藏与插入：把选中形状存为单页 .pptx 入库，或从库中打开并粘贴到当前页。
    /// </summary>
    public static class ShapeExporter
    {
        private static PowerPoint.Application App
        {
            get { return Globals.ThisAddIn.Application; }
        }

        // 说明（重要）：曾尝试"库文稿句柄池"（复用已打开的库 pptx 以省去每次约 700ms 的 Open），
        // 但跨多次操作的 Presentation 句柄复用出现 FatalExecutionEngineError(0xc0000005) 致命崩溃，
        // 故已移除：恢复为"每次打开→用→关闭"，优先保证稳定。


        // ---------------- 收藏 ----------------

        /// <summary>无打开的演示文稿时抛出友好异常。</summary>
        private static void EnsurePresentation()
        {
            try
            {
                if (App.Presentations.Count == 0)
                {
                    throw new InvalidOperationException("请先新建或打开一个演示文稿，再使用收藏功能。");
                }
            }
            catch (InvalidOperationException) { throw; }
            catch
            {
                throw new InvalidOperationException("请先新建或打开一个演示文稿，再使用收藏功能。");
            }
        }

        /// <summary>收藏选中形状（整组，存为 pptx + 缩略图 + 元数据）。返回收藏 id，失败返回 null。</summary>
        public static string CollectShapes(string name)
        {
            EnsurePresentation();
            var selection = App.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange == null || selection.ShapeRange.Count == 0)
            {
                return null;
            }
            string autoName = GetAutoName(selection.ShapeRange) + " " + DateTime.Now.ToString("MM-dd HH:mm");
            if (!ResolveItemName(name, autoName, out string resolved)) return null;   // 用户取消命名
            return CollectRange(selection.ShapeRange, resolved);
        }

        /// <summary>
        /// 收藏前的命名决策（AutoName 开关真正生效的地方）：
        /// - 调用方显式给了 name → 直接用；
        /// - 设置 AutoName=true → 用自动名（含时间戳）；
        /// - AutoName=false → 弹输入框让用户命名（留空回退自动名），取消则放弃收藏。
        /// </summary>
        private static bool ResolveItemName(string explicitName, string autoName, out string resolved)
        {
            resolved = autoName;
            if (!string.IsNullOrWhiteSpace(explicitName))
            {
                resolved = explicitName;
                return true;
            }
            if (ShapeLibraryService.LoadSettings().AutoName)
            {
                return true;   // resolved = autoName
            }
            var input = new InputDialog("命名收藏", "收藏名称（留空则用自动名）：", autoName);
            if (input.ShowDialog() != System.Windows.Forms.DialogResult.OK) return false;
            resolved = string.IsNullOrWhiteSpace(input.Value) ? autoName : input.Value.Trim();
            return true;
        }

        /// <summary>收藏当前页全部形状（B2）。</summary>
        /// <summary>收藏整页（整页快照）：以幻灯片为单位复制当前页（含背景/版式/全部内容），
        /// 存入库；之后插入可原样还原整页。返回收藏 id。</summary>
        public static string CollectWholePage()
        {
            EnsurePresentation();
            try
            {
                var pres = App.ActivePresentation;
                if (pres == null) return null;
                var slide = pres.Slides[App.ActiveWindow.View.Slide.SlideIndex];

                string autoName = "整页 " + DateTime.Now.ToString("MM-dd HH:mm");
                if (!ResolveItemName(null, autoName, out string resolved)) return null;   // 用户取消命名

                string id = ShapeLibraryService.NewId();
                ShapeLibraryService.CreateItemDir(id);
                PowerPoint.Presentation tempPres = null;
                var clip = ClipboardGuard.Snapshot();   // 兜底保护（仅在走剪贴板回退路径时才真正被覆盖）
                var swAll = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    tempPres = App.Presentations.Add(MsoTriState.msoFalse);
                    // 关键：沿用源页面的尺寸与宽高比。此前固定 ppSlideSizeOnScreen(4:3) 会把 16:9 模板
                    // 压成 4:3 存档，插回原文稿时页面变形、看起来"只是形状"。
                    try
                    {
                        tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeCustom;
                        tempPres.PageSetup.SlideWidth = pres.PageSetup.SlideWidth;
                        tempPres.PageSetup.SlideHeight = pres.PageSetup.SlideHeight;
                    }
                    catch { }

                    // 用"复制幻灯片 → 粘贴到临时文稿"（剪贴板）收藏：只有这条路径会把该页的
                    // **设计/版式绑定**一起带过来，收藏后的缩略图才与源页一致。
                    // 曾改用 Slides.InsertFromFile（免剪贴板），但它不带源页的版式绑定，PowerPoint 会把
                    // 插入页重绑到临时文稿的默认版式 → 占位符继承到线条，缩略图/插入结果就出现边框（已回退）。
                    slide.Copy();
                    System.Threading.Thread.Sleep(60);
                    {
                        var pr = tempPres.Slides.Paste();
                        if (pr == null || pr.Count == 0) throw new Exception("页面复制失败");
                    }
                    // 主题色固化：把整页里各形状的颜色按源页实际显示固化为显式 RGB，避免 A→B
                    try { BakeSlideColors(slide, tempPres.Slides[tempPres.Slides.Count]); } catch { }

                    // 删除新建文稿自带的空白页（InsertFromFile 路径下会残留一页）
                    while (tempPres.Slides.Count > 1)
                    {
                        bool removed = false;
                        for (int i = 1; i < tempPres.Slides.Count; i++)
                        {
                            try
                            {
                                // 删掉"没有内容且不是我们插入的那页"的空白页（新建文稿默认页）
                                if (tempPres.Slides[i].Shapes.Count == 0) { tempPres.Slides[i].Delete(); removed = true; break; }
                            }
                            catch { }
                        }
                        if (!removed) break;
                    }

                    // 注意：**不要**在这里删除"未使用版式"。
                    // 曾用 CleanUnusedLayouts 精简档（1.5MB→数百 KB），但其"哪些版式被使用"的判断
                    // （按 版式名+设计名 匹配）不可靠，会把本页真正使用的版式删掉，PowerPoint 随即把
                    // 这一页重绑到另一个带占位符的版式 → 占位符继承到线条，用户看到"形状自身明明是无
                    // 线条，却渲染出边框"。保真优先：保留源文件全部版式/母版。

                    // 设计/母版/版式改唯一名：避免插入时被目标文稿"按名字复用版式"而导致外观走样（多出边框/颜色变化）
                    try { MakeDesignUniqueForLibrary(tempPres, id); } catch { }

                    // 缩略图：按页面真实宽高比导出（避免拉伸变形）
                    string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                    try
                    {
                        int pw = 960;
                        int ph = (int)Math.Round(960.0 * pres.PageSetup.SlideHeight / Math.Max(1.0, pres.PageSetup.SlideWidth));
                        if (ph <= 0) ph = 540;
                        tempPres.Slides[tempPres.Slides.Count].Export(previewPath, "PNG", pw, ph);
                        if (System.IO.File.Exists(previewPath))
                            ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                    }
                    catch { }
                    finally { TryDelete(previewPath); }

                    // 保存库文件（单页含背景/版式）
                    string savePath = ShapeLibraryService.ShapeFile(id);
                    tempPres.SaveAs(savePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                    swAll.Stop();
                    try
                    {
                        long kb = System.IO.File.Exists(savePath) ? new System.IO.FileInfo(savePath).Length / 1024 : 0;
                        Log.Write("collect page timing: total=" + swAll.ElapsedMilliseconds + "ms file=" + kb + "KB");
                    }
                    catch { }

                    var meta = new LibraryMeta
                    {
                        Id = id,
                        Name = resolved,
                        Kind = "Shape",
                        CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        ShapeCount = 0,
                        SourceSlide = GetSourceInfo(),
                        Group = ShapeLibraryService.BrowseGroupForCollect,
                        Category = ShapeLibraryService.BrowseCategoryForCollect ?? "未分类",
                        SnapshotPage = true,
                        LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        SortOrder = Environment.TickCount,
                        Width = 0,
                        Height = 0
                    };
                    ShapeLibraryService.SaveMeta(meta);
                    return id;
                }
                finally
                {
                    ClipboardGuard.Restore(clip);
                    if (tempPres != null) { try { tempPres.Close(); } catch { } }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("收藏整页失败：" + ex.Message, "素材库");
                return null;
            }
        }

        /// <summary>清理演示文稿中"未被任何幻灯片使用"的自定义版式（每个母版至少保留 1 个）。
        /// 用于插入整页快照后，避免源文件版式在目标文稿里无限累积。返回删除数量。</summary>
        private static int CleanUnusedLayouts(PowerPoint.Presentation pres)
        {
            int removed = 0;
            try
            {
                // 已使用版式名集合（同母版内版式名唯一；跨母版以 母版名+版式名 区分）
                var used = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    try
                    {
                        var layout = pres.Slides[i].CustomLayout;
                        if (layout != null) used.Add((layout.Name ?? "") + "@" + (layout.Design != null ? layout.Design.Name : ""));
                    }
                    catch { }
                }
                for (int d = 1; d <= pres.Designs.Count; d++)
                {
                    PowerPoint.Design design = null;
                    try { design = pres.Designs[d]; } catch { continue; }
                    PowerPoint.Master master = null;
                    try { master = design.SlideMaster; } catch { continue; }
                    if (master == null) continue;
                    int layoutCount = 0;
                    try { layoutCount = master.CustomLayouts.Count; } catch { }
                    // 倒序删除，保留第一个（默认版式）
                    for (int li = layoutCount; li >= 2; li--)
                    {
                        try
                        {
                            var layout = master.CustomLayouts[li];
                            string key = (layout.Name ?? "") + "@" + (design.Name ?? "");
                            if (used.Contains(key)) continue;
                            layout.Delete();
                            removed++;
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return removed;
        }

        // ---------------- 主题色固化（修"收藏后颜色 A→B"） ----------------
        // PowerPoint 中颜色常为主题色（Accent 1 等）。把形状/整页粘进新建的空白文稿时，PowerPoint 按
        // 主题名匹配到空白文稿的默认 Office 主题，主题色被重新解析 → 颜色变了（预览与插入都变）。
        // 收藏时按"源形状当前实际显示的颜色"把颜色固化成显式 RGB，之后预览与插入到任何文稿都保持原色。

        private static void BakeRangeColors(PowerPoint.ShapeRange src, PowerPoint.ShapeRange dst)
        {
            try
            {
                int n = Math.Min(src.Count, dst.Count);
                for (int i = 1; i <= n; i++)
                {
                    try { BakeShapeColors(src[i], dst[i]); } catch { }
                }
            }
            catch { }
        }

        private static void BakeSlideColors(PowerPoint.Slide src, PowerPoint.Slide dst)
        {
            try
            {
                int n = Math.Min(src.Shapes.Count, dst.Shapes.Count);
                for (int i = 1; i <= n; i++)
                {
                    try { BakeShapeColors(src.Shapes[i], dst.Shapes[i]); } catch { }
                }
            }
            catch { }

            // 背景：仅当来源页"显式设置了背景"（不跟随母版）且为纯色时才固化到本页；
            // 跟随母版时不动（改它会连带影响母版，且这类跨对象操作风险高）
            try
            {
                if (src.FollowMasterBackground == MsoTriState.msoFalse)
                {
                    var sFill = src.Background.Fill;
                    if ((int)sFill.Type == 1)
                    {
                        int rgb = sFill.ForeColor.RGB;
                        dst.FollowMasterBackground = MsoTriState.msoFalse;
                        dst.Background.Fill.Solid();
                        dst.Background.Fill.ForeColor.RGB = rgb;
                    }
                }
            }
            catch { }
        }

        private static void BakeShapeColors(PowerPoint.Shape src, PowerPoint.Shape dst)
        {
            // 组合：按顺序递归子形状
            try
            {
                if (src.Type == MsoShapeType.msoGroup && dst.Type == MsoShapeType.msoGroup)
                {
                    int n = Math.Min(src.GroupItems.Count, dst.GroupItems.Count);
                    for (int i = 1; i <= n; i++)
                    {
                        try { BakeShapeColors(src.GroupItems[i], dst.GroupItems[i]); } catch { }
                    }
                }
            }
            catch { }

            // 填充：纯色 / 渐变（含各停靠点）
            try
            {
                int st = (int)src.Fill.Type;
                if (st == 1)                    // msoFillSolid
                {
                    dst.Fill.Solid();
                    dst.Fill.ForeColor.RGB = src.Fill.ForeColor.RGB;
                }
                else if (st == 3)               // msoFillGradient
                {
                    try { dst.Fill.ForeColor.RGB = src.Fill.ForeColor.RGB; } catch { }
                    try { dst.Fill.BackColor.RGB = src.Fill.BackColor.RGB; } catch { }
                    try
                    {
                        var ss = src.Fill.GradientStops;
                        var ds = dst.Fill.GradientStops;
                        int k = Math.Min(ss.Count, ds.Count);
                        for (int i = 1; i <= k; i++)
                        {
                            try { ds[i].Color.RGB = ss[i].Color.RGB; } catch { }
                        }
                    }
                    catch { }
                }
                else if (st == 4 || st == 5)    // 纹理 / 背景填充
                {
                    try { dst.Fill.ForeColor.RGB = src.Fill.ForeColor.RGB; } catch { }
                }
            }
            catch { }

            // 线条
            try
            {
                if (src.Line.Visible == MsoTriState.msoTrue)
                    dst.Line.ForeColor.RGB = src.Line.ForeColor.RGB;
            }
            catch { }

            // 文字颜色（整段；注意：逐 Run 的多色文本仍可能随主题重映射，属已知限制）
            try
            {
                var stf = src.TextFrame2;
                var dtf = dst.TextFrame2;
                if (stf != null && dtf != null && stf.HasText == MsoTriState.msoTrue && dtf.HasText == MsoTriState.msoTrue)
                {
                    try { dtf.TextRange.Font.Fill.ForeColor.RGB = stf.TextRange.Font.Fill.ForeColor.RGB; } catch { }
                }
            }
            catch { }

            // 阴影 / 三维
            try { dst.Shadow.ForeColor.RGB = src.Shadow.ForeColor.RGB; } catch { }
            try { dst.ThreeD.ExtrusionColor.RGB = src.ThreeD.ExtrusionColor.RGB; } catch { }
            try { dst.ThreeD.ContourColor.RGB = src.ThreeD.ContourColor.RGB; } catch { }
        }

        // ---------------- 设计/版式改名（修"插入后外观走样"） ----------------
        // PowerPoint 在把幻灯片插入到别的文稿时，会按 **设计/版式名字** 去匹配目标文稿里已有的版式；
        // 名字相同就直接复用目标版式 → 占位符的线条/填充/字体/颜色随目标版式变化（用户看到"多了边框"、
        // "颜色变了"）。收藏时把本条目携带的设计/母版/版式改成唯一名字，即可迫使目标文稿导入我们自己的
        // 设计，从而保持原样。
        private static void MakeDesignUniqueForLibrary(PowerPoint.Presentation pres, string id)
        {
            try
            {
                string tag = "Lib" + id;
                for (int d = 1; d <= pres.Designs.Count; d++)
                {
                    PowerPoint.Design design = null;
                    try { design = pres.Designs[d]; } catch { continue; }
                    if (design == null) continue;
                    try { design.Name = tag + "-D" + d; } catch { }
                    PowerPoint.Master master = null;
                    try { master = design.SlideMaster; } catch { }
                    if (master == null) continue;
                    try { master.Name = tag + "-M" + d; } catch { }
                    try
                    {
                        for (int li = 1; li <= master.CustomLayouts.Count; li++)
                        {
                            try { master.CustomLayouts[li].Name = tag + "-L" + d + "-" + li; } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>插入整页快照：以幻灯片级复制粘贴还原整页（保留背景/版式）。
        /// position: "after"=当前页之后(默认) / "end"=文稿末尾 / "newDoc"=新建演示文稿打开。
        /// 失败原因写入 error（Web 桥 Toast 呈现）。</summary>
        public static bool InsertPageSnapshot(string id, string position, out string error)
        {
            error = null;
            try
            {
                string pos = string.IsNullOrWhiteSpace(position) ? "after" : position;
                bool newDoc = pos == "newDoc";
                if (!newDoc && App.Presentations.Count == 0)
                {
                    error = "请先新建或打开演示文稿，再插入整页。";
                    return false;
                }
                var meta = ShapeLibraryService.LoadMeta(id);
                if (meta == null || !meta.SnapshotPage)
                {
                    error = "这不是整页快照。";
                    return false;
                }
                string shapePath = ShapeLibraryService.ShapeFile(id);
                if (!System.IO.File.Exists(shapePath))
                {
                    error = "整页文件缺失。";
                    return false;
                }

                var swPage = System.Diagnostics.Stopwatch.StartNew();
                // 打开库文件 → 复制该页（幻灯片级）→ 粘贴到目标位置。
                // 只有"复制幻灯片"这条路径会把该页的**设计/版式绑定**一起带到目标文稿；
                // 曾改用 Slides.InsertFromFile（免剪贴板），但它不带版式绑定，PowerPoint 会把插入页重绑到
                // 目标文稿的默认版式 → 占位符继承线条/颜色，插入后多出边框（已回退）。
                PowerPoint.Presentation libPres = null;
                var clip = ClipboardGuard.Snapshot();
                try
                {
                    libPres = App.Presentations.Open(shapePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    if (libPres.Slides.Count == 0) { error = "整页内容为空。"; return false; }
                    libPres.Slides[1].Copy();
                    System.Threading.Thread.Sleep(60);

                    if (newDoc)
                    {
                        var np = App.Presentations.Add(MsoTriState.msoTrue);
                        var pastedNew = np.Slides.Paste();
                        if (pastedNew == null || pastedNew.Count == 0) { error = "新建文稿粘贴失败。"; return false; }
                        ShapeLibraryService.TouchUsed(id);
                        swPage.Stop();
                        Log.Write("insert page timing: total=" + swPage.ElapsedMilliseconds + "ms pos=newDoc");
                        return true;
                    }

                    var pres = App.ActivePresentation;
                    int insertAt = pos == "end" ? pres.Slides.Count + 1 : App.ActiveWindow.View.Slide.SlideIndex + 1;
                    if (insertAt > pres.Slides.Count + 1) insertAt = pres.Slides.Count + 1;
                    if (insertAt < 1) insertAt = 1;

                    var inserted = pres.Slides.Paste(insertAt);
                    if (inserted == null || inserted.Count == 0) { error = "整页粘贴失败。"; return false; }

                    // 注意：这里**不再**执行"清理未使用版式"（判断不可靠，会改坏版式绑定）。
                    try
                    {
                        if (ShapeLibraryService.LoadSettings().CleanLayoutsOnInsertPage)
                            Log.Write("cleanLayoutsOnInsertPage=on 但已停用（会改坏版式绑定），本次跳过");
                    }
                    catch { }

                    ShapeLibraryService.TouchUsed(id);
                    swPage.Stop();
                    Log.Write("insert page timing: total=" + swPage.ElapsedMilliseconds + "ms pos=" + pos);
                    return true;
                }
                finally
                {
                    ClipboardGuard.Restore(clip);
                    if (libPres != null) { try { libPres.Close(); } catch { } }
                }
            }
            catch (Exception ex)
            {
                error = "插入整页失败：" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 批量收藏整份演示文稿：每页所有非占位形状打包为一个收藏（跳过空页）。
        /// 使用安静模式：不去重弹框、不逐个询问命名（命名统一为 "第N页 {自动名}"）。
        /// 返回成功收藏数。
        /// </summary>
        public static int CollectAllSlides()
        {
            EnsurePresentation();
            var pres = App.ActivePresentation;
            if (pres == null) return 0;
            int okCount = 0;
            try
            {
                for (int s = 1; s <= pres.Slides.Count; s++)
                {
                    var slide = pres.Slides[s];
                    var idx = new System.Collections.Generic.List<int>();
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                            if (!IsEmptyPlaceholder(slide.Shapes[i])) idx.Add(i);
                    }
                    if (idx.Count == 0) continue;
                    var range = slide.Shapes.Range(idx.ToArray());
                    string nm = "第" + s + "页 " + GetAutoName(range) + " " + DateTime.Now.ToString("MM-dd HH:mm");
                    string id = CollectRangeQuiet(range, nm);
                    if (id != null) okCount++;
                }
                return okCount;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("批量收藏失败：" + ex.Message, "素材库");
                return okCount;
            }
        }

        /// <summary>
        /// 从剪贴板收藏：不要求预先选中形状。将剪贴板内容粘贴到临时空白页后整组收藏。
        /// 支持 PPT 形状、图片等（经 PowerPoint 剪贴板粘贴）。返回收藏 id，无可粘贴内容返回 null。
        /// </summary>
        public static string CollectClipboard()
        {
            EnsurePresentation();
            PowerPoint.Presentation tempPres = null;
            try
            {
                // 剪贴板无内容直接放弃
                if (!System.Windows.Forms.Clipboard.ContainsData("PowerPoint Shapedata")
                    && !System.Windows.Forms.Clipboard.ContainsImage()
                    && !System.Windows.Forms.Clipboard.ContainsFileDropList()
                    && !System.Windows.Forms.Clipboard.ContainsText())
                {
                    System.Windows.Forms.MessageBox.Show("剪贴板中没有可收藏的内容。", "素材库");
                    return null;
                }

                tempPres = App.Presentations.Add(MsoTriState.msoFalse);
                tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                var slide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                // 剪贴板内容直接粘贴到临时页（Paste 返回 ShapeRange；带重试应对 Office 异步剪贴板）
                PowerPoint.ShapeRange pasted = null;
                Exception lastErr = null;
                for (int attempt = 1; attempt <= 5; attempt++)
                {
                    try
                    {
                        pasted = slide.Shapes.Paste();
                        if (pasted != null && pasted.Count > 0) break;
                    }
                    catch (Exception ex)
                    {
                        lastErr = ex;
                        System.Threading.Thread.Sleep(120 * attempt);
                    }
                }
                if (pasted == null || pasted.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "剪贴板内容无法粘贴为形状" + (lastErr != null ? "：" + lastErr.Message : "") +
                        "\n请先在 PowerPoint 中复制形状/图片后再试。", "素材库");
                    return null;
                }
                return CollectRangeQuiet(pasted, null);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("剪贴板收藏失败：" + ex.Message, "素材库");
                return null;
            }
            finally
            {
                if (tempPres != null)
                {
                    try { tempPres.Close(); } catch { }
                }
            }
        }

        /// <summary>把形状(含组合内子形状)的文本替换为占位提示文字。
        /// 保留首字符的字体名/字号/加粗/斜体/颜色(以 PPT 原生 int RGB 直读直写，避免 ARGB/BGR 错乱)/段落对齐。</summary>
        private static void ReplaceShapeTextWithPlaceholder(PowerPoint.Shape shape, string placeholder)
        {
            try
            {
                if (shape.Type == MsoShapeType.msoGroup && shape.GroupItems != null)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        ReplaceShapeTextWithPlaceholder(shape.GroupItems[i], placeholder);
                    }
                    return;
                }
                var tf = shape.TextFrame2;
                if (tf == null || tf.HasText != MsoTriState.msoTrue) return;
                var tr = tf.TextRange;   // TextFrame2.TextRange 是 Office.TextRange2
                if (tr == null || tr.Length == 0) return;
                string oldText = "";
                try { oldText = tr.Text; } catch { }
                if (oldText == placeholder) return;

                // 抓首字符格式（用 PPT 原生 int RGB，直读直写避免 BGR/ARGB 转换错乱）
                string fontName = null; float fontSize = 0; int bold = -1; int italic = -1;
                int colorRgb = -1; int align = -1;
                try { fontName = tr.Font.Name; } catch { }
                try { fontSize = tr.Font.Size; } catch { }
                try { bold = tr.Font.Bold == MsoTriState.msoTrue ? 1 : (tr.Font.Bold == MsoTriState.msoFalse ? 0 : -1); } catch { }
                try { italic = tr.Font.Italic == MsoTriState.msoTrue ? 1 : (tr.Font.Italic == MsoTriState.msoFalse ? 0 : -1); } catch { }
                try { colorRgb = tr.Font.Fill.ForeColor.RGB; } catch { }
                try { align = (int)tr.ParagraphFormat.Alignment; } catch { }

                tr.Text = placeholder;

                // 回刷格式（占位文字继承原观感）
                try { if (!string.IsNullOrEmpty(fontName)) tr.Font.Name = fontName; } catch { }
                try { if (fontSize > 0) tr.Font.Size = fontSize; } catch { }
                try { if (bold >= 0) tr.Font.Bold = bold == 1 ? MsoTriState.msoTrue : MsoTriState.msoFalse; } catch { }
                try { if (italic >= 0) tr.Font.Italic = italic == 1 ? MsoTriState.msoTrue : MsoTriState.msoFalse; } catch { }
                try { if (colorRgb >= 0) tr.Font.Fill.ForeColor.RGB = colorRgb; } catch { }
                try { if (align >= 0) tr.ParagraphFormat.Alignment = (Microsoft.Office.Core.MsoParagraphAlignment)align; } catch { }
            }
            catch { /* 单个形状文字替换失败不阻断收藏 */ }
        }

        /// <summary>收藏任意 ShapeRange 的核心实现。</summary>
        public static string CollectRange(PowerPoint.ShapeRange range, string name)
        {
            return CollectRangeCore(range, name, false);
        }

        /// <summary>安静模式收藏（批量入库用）：不弹去重确认，不做命名询问，缩略图失败静默。</summary>
        public static string CollectRangeQuiet(PowerPoint.ShapeRange range, string name)
        {
            return CollectRangeCore(range, name, true);
        }

        private static string CollectRangeCore(PowerPoint.ShapeRange range, string name, bool quiet)
        {
            if (range == null || range.Count == 0) return null;

            PowerPoint.Presentation tempPres = null;
            string id = ShapeLibraryService.NewId();

            try
            {
                // 去重提示（quiet 模式跳过）
                if (!quiet)
                {
                    try
                    {
                        string dup = ShapeLibraryService.FindDuplicate(range);
                        if (dup != null)
                        {
                            if (System.Windows.Forms.MessageBox.Show(
                                "检测到可能重复的收藏：「" + dup + "」。仍要收藏吗？",
                                "素材库", System.Windows.Forms.MessageBoxButtons.YesNo,
                                System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
                            {
                                return null;
                            }
                        }
                    }
                    catch { }
                }

                // 0. 创建收藏目录
                ShapeLibraryService.CreateItemDir(id);

                // 2. 新建隐藏演示文稿，粘贴形状
                tempPres = App.Presentations.Add(MsoTriState.msoFalse);
                tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                var slide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                var clip = ClipboardGuard.Snapshot();   // 收藏也保护剪贴板（Copy 会覆盖）
                PowerPoint.ShapeRange pasted;
                try
                {
                    pasted = CopyPasteWithRetry(() => range.Copy(), slide.Shapes);
                }
                finally
                {
                    ClipboardGuard.Restore(clip);
                }
                // 主题色固化：按源形状实际显示的颜色写成显式 RGB，避免收藏后颜色随主题变化（A→B）
                try { BakeRangeColors(range, pasted); } catch { }
                // 归一化：把整组移动到 (0,0)，保证插入时坐标可控
                float minLeft = float.MaxValue, minTop = float.MaxValue;
                foreach (PowerPoint.Shape s in pasted)
                {
                    minLeft = Math.Min(minLeft, (float)s.Left);
                    minTop = Math.Min(minTop, (float)s.Top);
                }
                foreach (PowerPoint.Shape s in pasted)
                {
                    s.Left -= minLeft;
                    s.Top -= minTop;
                }

                // 设置「收藏时替换文字为占位提示」：把库副本内文字换成占位文本（不改源画布），
                // 生成的缩略图也显示占位文字，插入后用户直接输入真实内容即可。
                if (ShapeLibraryService.LoadSettings().CollectPlaceholderText)
                {
                    foreach (PowerPoint.Shape s in pasted)
                    {
                        ReplaceShapeTextWithPlaceholder(s, "双击编辑文字");
                    }
                }

                // 1. 生成缩略图：从"库文稿中已归一化的形状"导出（与插入到用户页面时的内容一致），
                //    不再直接从源页面导出——避免源页面排版/页面缩放导致预览文字与形状错位
                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                try
                {
                    pasted.Export(previewPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                }
                catch { /* 缩略图失败不阻断 */ }
                finally { TryDelete(previewPath); }

                // 3. 保存为库文件
                string savePath = ShapeLibraryService.ShapeFile(id);
                tempPres.SaveAs(savePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                // 4. 元数据
                float tw = 0, th = 0;
                try
                {
                    foreach (PowerPoint.Shape s in range)
                    {
                        tw += (float)s.Width;
                        th += (float)s.Height;
                    }
                }
                catch { }
                var meta = new LibraryMeta
                {
                    Id = id,
                    Name = string.IsNullOrWhiteSpace(name) ? GetAutoName(range) + " " + DateTime.Now.ToString("MM-dd HH:mm") : name,
                    Kind = "Shape",
                    CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ShapeCount = range.Count,
                    SourceSlide = GetSourceInfo(),
                    // 收藏落点 = 面板当前浏览的 组/分类（用户正打开的视图）；未明确时回落 我的/未分类
                    Group = ShapeLibraryService.BrowseGroupForCollect,
                    Category = ShapeLibraryService.BrowseCategoryForCollect ?? "未分类",
                    ColorHex = GetPrimaryColorHex(range),
                    LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SortOrder = Environment.TickCount,
                    Width = tw,
                    Height = th
                };
                ShapeLibraryService.SaveMeta(meta);

                return id;
            }
            catch (Exception ex)
            {
                if (!quiet)
                {
                    System.Windows.Forms.MessageBox.Show("收藏失败：" + ex.Message, "素材库");
                }
                return null;
            }
            finally
            {
                if (tempPres != null)
                {
                    try { tempPres.Close(); } catch { }
                }
            }
        }

        /// <summary>收藏选中形状的样式（存 JSON + 缩略图 + 元数据）。</summary>
        public static string CollectStyle(string name)
        {
            EnsurePresentation();
            var selection = App.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange == null || selection.ShapeRange.Count == 0)
            {
                return null;
            }
            string autoName = "样式 " + DateTime.Now.ToString("MM-dd HH:mm");
            if (!ResolveItemName(name, autoName, out string resolved)) return null;   // 用户取消命名
            return CollectStyle(selection.ShapeRange, resolved);
        }

        /// <summary>样式收藏核心实现（命名已解析）。</summary>
        private static string CollectStyle(PowerPoint.ShapeRange range, string name)
        {
            string id = ShapeLibraryService.NewId();
            try
            {
                ShapeLibraryService.CreateItemDir(id);
                var preset = ShapeStyleExtractor.Extract(range, name);
                ShapeLibraryService.SaveStyleJson(id, ShapeStyleExtractor.ToJson(preset));

                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                try
                {
                    // 自然尺寸导出，避免固定宽高导致文字/形状比例错位
                    range.Export(previewPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                }
                catch { }
                finally { TryDelete(previewPath); }

                var meta = new LibraryMeta
                {
                    Id = id,
                    Name = string.IsNullOrWhiteSpace(name) ? "样式 " + DateTime.Now.ToString("MM-dd HH:mm") : name,
                    Kind = "Style",
                    CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ShapeCount = range.Count,
                    SourceSlide = GetSourceInfo(),
                    Group = ShapeLibraryService.BrowseGroupForCollect,
                    Category = ShapeLibraryService.BrowseCategoryForCollect ?? "未分类",
                    ColorHex = GetPrimaryColorHex(range),
                    LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SortOrder = Environment.TickCount,
                    SourceAutoType = GetAutoType(range[1])
                };
                ShapeLibraryService.SaveMeta(meta);
                return id;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("收藏样式失败：" + ex.Message, "素材库");
                return null;
            }
        }

        // ---------------- 插入 / 应用 ----------------

        /// <summary>把库中形状插入到当前页中心。返回是否成功（失败弹框提示）。</summary>
        public static bool InsertShape(string id)
        {
            return InsertShape(id, true, out _);
        }

        /// <summary>
        /// 插入核心。showError=true 时失败弹 MessageBox（Ribbon 直调路径）；
        /// false 时静默返回 false 并把错误原因写入 error（Web 桥路径，前端以 Toast 呈现）。
        /// </summary>
        public static bool InsertShape(string id, bool showError, out string error)
        {
            error = null;
            // 防御：整页快照条目必须按"页"插入。若被当成普通形状插入，就会把页内形状粘到当前页上，
            // 看起来"收藏整页其实只收藏了形状"。这里强制改走整页插入。
            try
            {
                var probe = ShapeLibraryService.LoadMeta(id);
                if (probe != null && probe.SnapshotPage)
                {
                    return InsertPageSnapshot(id, "after", out error);
                }
            }
            catch { }

            try
            {
                EnsurePresentation();
            }
            catch (InvalidOperationException ex)
            {
                error = ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            string path = ShapeLibraryService.ShapeFile(id);
            if (!System.IO.File.Exists(path))
            {
                error = "收藏文件不存在，可能已被移动或删除。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }

            PowerPoint.Presentation libPres = null;
            try
            {
                var swIns = System.Diagnostics.Stopwatch.StartNew();
                libPres = App.Presentations.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

                var slide = libPres.Slides[1];
                if (slide.Shapes.Count == 0)
                {
                    error = "该收藏没有可插入的形状。";
                    return false;
                }

                // 构造整组 ShapeRange 一次性复制（Range 参数为位置索引 1..Count）
                int[] ids = new int[slide.Shapes.Count];
                for (int i = 1; i <= slide.Shapes.Count; i++) ids[i - 1] = i;
                var range = slide.Shapes.Range(ids);

                PowerPoint.Shapes targetShapes = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex].Shapes;
                var clip = ClipboardGuard.Snapshot();   // 保护用户剪贴板（Copy 会覆盖它）
                PowerPoint.ShapeRange pasted = null;
                try
                {
                    pasted = CopyPasteWithRetry(() => range.Copy(), targetShapes);
                }
                finally
                {
                    ClipboardGuard.Restore(clip);        // 尽快恢复，避免残留我们的内容
                }

                // 粘贴到页面中心
                float sw = (float)App.ActivePresentation.PageSetup.SlideWidth;
                float sh = (float)App.ActivePresentation.PageSetup.SlideHeight;
                float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
                foreach (PowerPoint.Shape s in pasted)
                {
                    minL = Math.Min(minL, (float)s.Left);
                    minT = Math.Min(minT, (float)s.Top);
                    maxR = Math.Max(maxR, (float)(s.Left + s.Width));
                    maxB = Math.Max(maxB, (float)(s.Top + s.Height));
                }

                // 超大形状等比缩小到画布 90% 以内
                float groupW = maxR - minL, groupH = maxB - minT;
                float maxW = sw * 0.9f, maxH = sh * 0.9f;
                if (groupW > maxW || groupH > maxH)
                {
                    float scale = Math.Min(maxW / groupW, maxH / groupH);
                    bool scaled = true;
                    try
                    {
                        pasted.ScaleWidth(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                        pasted.ScaleHeight(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                    }
                    catch { scaled = false; }   // 缩放失败：保持原位置（下方不再居中），防止按未缩放尺寸计算而偏出画布
                    // 缩放后重新计算边界
                    minL = float.MaxValue; minT = float.MaxValue; maxR = float.MinValue; maxB = float.MinValue;
                    foreach (PowerPoint.Shape s in pasted)
                    {
                        minL = Math.Min(minL, (float)s.Left);
                        minT = Math.Min(minT, (float)s.Top);
                        maxR = Math.Max(maxR, (float)(s.Left + s.Width));
                        maxB = Math.Max(maxB, (float)(s.Top + s.Height));
                    }
                    if (!scaled)
                    {
                        // 缩放失败：不居中（dx/dy 会基于失真尺寸），仅记录并保持原坐标
                        try { pasted.Select(); } catch { }
                        ShapeLibraryService.TouchUsed(id);
                        return true;
                    }
                }

                float dx = (sw - (maxR - minL)) / 2f - minL;
                float dy = (sh - (maxB - minT)) / 2f - minT;
                bool center = ShapeLibraryService.LoadSettings().InsertPosition != "original";
                foreach (PowerPoint.Shape s in pasted)
                {
                    if (center)
                    {
                        s.Left += dx;
                        s.Top += dy;
                    }
                }

                // 记录最近使用
                ShapeLibraryService.TouchUsed(id);

                // 插入后自动选中，方便继续调整
                try { pasted.Select(); } catch { }
                swIns.Stop();
                if (swIns.ElapsedMilliseconds > 300)
                    Log.Write("insert timing: total=" + swIns.ElapsedMilliseconds + "ms shapes=" + slide.Shapes.Count);
                return true;
            }
            catch (Exception ex)
            {
                error = "插入失败：" + ex.Message;
                try
                {
                    string viewInfo = "";
                    try
                    {
                        var v = App.ActiveWindow != null ? App.ActiveWindow.View : null;
                        viewInfo = v != null ? ("view=" + v.Type) : "no-active-window";
                    }
                    catch { viewInfo = "view-unknown"; }
                    Log.Write("InsertShape failed id=" + id + " " + viewInfo + " :: " + ex.Message);
                }
                catch { }
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            finally
            {
                if (libPres != null) { try { libPres.Close(); } catch { } }
            }
        }

        /// <summary>把库中样式应用到当前选中形状。失败弹框提示。</summary>
        public static bool ApplyStyle(string id)
        {
            return ApplyStyle(id, true, out _);
        }

        /// <summary>静默版：Web 桥路径，错误原因写入 error 由前端 Toast 呈现。</summary>
        public static bool ApplyStyle(string id, bool showError, out string error)
        {
            error = null;
            try
            {
                EnsurePresentation();
            }
            catch (InvalidOperationException ex)
            {
                error = ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            string json = ShapeLibraryService.LoadStyleJson(id);
            if (json == null)
            {
                error = "样式文件不存在。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            var selection = App.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange == null || selection.ShapeRange.Count == 0)
            {
                error = "请先在页面上选中要应用样式的形状。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            var preset = ShapeStyleExtractor.FromJson(json);
            if (preset == null)
            {
                error = "样式解析失败。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            try
            {
                AdaptPresetToTheme(preset);
                ShapeStyleExtractor.Apply(selection.ShapeRange, preset);
                return true;
            }
            catch (Exception ex)
            {
                error = "应用样式失败：" + ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
        }

        /// <summary>把样式应用到当前页所有与来源同类型的形状（D3）。失败弹框提示。</summary>
        public static bool ApplyStyleToSameType(string id)
        {
            return ApplyStyleToSameType(id, true, out _);
        }

        /// <summary>静默版：Web 桥路径。</summary>
        public static bool ApplyStyleToSameType(string id, bool showError, out string error)
        {
            error = null;
            var meta = ShapeLibraryService.LoadMeta(id);
            string json = ShapeLibraryService.LoadStyleJson(id);
            if (meta == null || json == null || meta.Kind != "Style")
            {
                error = "这不是样式收藏。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            var preset = ShapeStyleExtractor.FromJson(json);
            if (preset == null)
            {
                error = "样式解析失败。";
                return false;
            }
            if (meta.SourceAutoType <= 0)
            {
                error = "该样式未记录来源形状类型，无法批量应用。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            try
            {
                var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                int applied = 0;
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var s = slide.Shapes[i];
                    try
                    {
                        if ((int)s.AutoShapeType == meta.SourceAutoType)
                        {
                            var range = slide.Shapes.Range(new[] { i });
                            AdaptPresetToTheme(preset);
                            ShapeStyleExtractor.Apply(range, preset);
                            applied++;
                        }
                    }
                    catch { }
                }
                if (applied == 0)
                {
                    error = "本页没有找到同类型的形状。";
                    if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                    return false;
                }
                if (showError)
                {
                    System.Windows.Forms.MessageBox.Show("已应用到 " + applied + " 个同类型形状。", "素材库");
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "批量应用失败：" + ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
        }

        /// <summary>把样式应用到整份演示文稿中所有与来源同类型的形状。失败弹框提示。</summary>
        public static bool ApplyStyleToSameTypeAllSlides(string id)
        {
            return ApplyStyleToSameTypeAllSlides(id, true, out _);
        }

        /// <summary>静默版：Web 桥路径。</summary>
        public static bool ApplyStyleToSameTypeAllSlides(string id, bool showError, out string error)
        {
            error = null;
            var meta = ShapeLibraryService.LoadMeta(id);
            string json = ShapeLibraryService.LoadStyleJson(id);
            if (meta == null || json == null || meta.Kind != "Style")
            {
                error = "这不是样式收藏。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            var preset = ShapeStyleExtractor.FromJson(json);
            if (preset == null)
            {
                error = "样式解析失败。";
                return false;
            }
            if (meta.SourceAutoType <= 0)
            {
                error = "该样式未记录来源形状类型，无法批量应用。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
            try
            {
                var pres = App.ActivePresentation;
                int applied = 0;
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var s = slide.Shapes[i];
                        try
                        {
                            if ((int)s.AutoShapeType == meta.SourceAutoType)
                            {
                                var range = slide.Shapes.Range(new[] { i });
                                AdaptPresetToTheme(preset);
                                ShapeStyleExtractor.Apply(range, preset);
                                applied++;
                            }
                        }
                        catch { }
                    }
                }
                if (applied == 0)
                {
                    error = "整份演示文稿中没有找到同类型的形状。";
                    if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                    return false;
                }
                if (showError)
                {
                    System.Windows.Forms.MessageBox.Show("已在 " + applied + " 个同类型形状上应用样式。", "素材库");
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "批量应用失败：" + ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "素材库");
                return false;
            }
        }

        /// <summary>D1：设置开启时把样式主色映射到当前文档主题强调色。</summary>
        private static void AdaptPresetToTheme(StylePreset preset)
        {
            try
            {
                if (!ShapeLibraryService.LoadSettings().ThemeAdaptStyle) return;
                var slide = App.ActivePresentation.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                var accent = slide.ThemeColorScheme(MsoThemeColorIndex.msoThemeColorAccent1).RGB;
                string hex = ShapeStyleExtractor.RgbToHex(accent);
                if (preset.FillColor != null) preset.FillColor = hex;
                if (preset.LineColor != null) preset.LineColor = hex;
                if (preset.FontColor != null) preset.FontColor = hex;
                if (preset.GlowColor != null) preset.GlowColor = hex;
            }
            catch { }
        }

        // ---------------- 内置预设素材库（B1） ----------------

        private static string BuiltinMarker
        {
            get { return System.IO.Path.Combine(ShapeLibraryService.RootDir, ".builtin_installed"); }
        }

        /// <summary>首次运行时安装内置预设素材（15 个常用形状组合）。每类使用独立隐藏文稿，避免嵌套操作。</summary>
        public static void InstallBuiltinLibrary()
        {
            try
            {
                if (System.IO.File.Exists(BuiltinMarker)) return;

                var specs = new System.Collections.Generic.List<Tuple<string, string>>();
                specs.Add(Tuple.Create("圆角卡片", "card"));
                specs.Add(Tuple.Create("渐变卡片", "gradientCard"));
                specs.Add(Tuple.Create("右箭头", "arrow"));
                specs.Add(Tuple.Create("流程步骤框", "process"));
                specs.Add(Tuple.Create("判定菱形", "decision"));
                specs.Add(Tuple.Create("标题副标题", "titleSub"));
                specs.Add(Tuple.Create("圆形图标底座", "iconCircle"));
                specs.Add(Tuple.Create("标签徽章", "badge"));
                specs.Add(Tuple.Create("分隔线", "divider"));
                specs.Add(Tuple.Create("引用块", "quote"));
                specs.Add(Tuple.Create("编号步骤", "stepNumber"));
                specs.Add(Tuple.Create("数据指标卡", "metric"));
                specs.Add(Tuple.Create("装饰圆环", "ring"));
                specs.Add(Tuple.Create("要点列表", "bulletList"));
                specs.Add(Tuple.Create("侧边栏条", "sidebar"));

                foreach (var spec in specs)
                {
                    try
                    {
                        PowerPoint.Presentation pres = null;
                        try
                        {
                            pres = App.Presentations.Add(MsoTriState.msoFalse);
                            pres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                            var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            DrawBuiltinShape(slide, spec.Item2);
                            if (slide.Shapes.Count == 0) continue;

                            string id = ShapeLibraryService.NewId();
                            ShapeLibraryService.CreateItemDir(id);

                            // 缩略图
                            try
                            {
                                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                                slide.Export(previewPath, "PNG", 320, 240);
                                ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                                TryDelete(previewPath);
                            }
                            catch { }

                            // 保存形状
                            pres.SaveAs(ShapeLibraryService.ShapeFile(id), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                            var meta = new LibraryMeta
                            {
                                Id = id,
                                Name = "内置·" + spec.Item1,
                                Kind = "Shape",
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                ShapeCount = slide.Shapes.Count,
                                Category = "内置素材",
                                LastUsedAt = "",
                                SortOrder = Environment.TickCount
                            };
                            ShapeLibraryService.SaveMeta(meta);
                        }
                        finally
                        {
                            if (pres != null) { try { pres.Close(); } catch { } }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Builtin:" + spec.Item1);
                    }
                }
                System.IO.File.WriteAllText(BuiltinMarker, DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Log.Error(ex, "InstallBuiltinLibrary");
            }
        }

        /// <summary>
        /// 一次性迁移：把旧分类(内置素材/学术模板)的条目归入对应一级组。只执行一次(marker)，
        /// 避免每次启动覆盖用户手动分组。
        /// </summary>
        public static void EnsureGroupAssignments()
        {
            try
            {
                string marker = System.IO.Path.Combine(ShapeLibraryService.RootDir, ".groups_migrated");
                if (System.IO.File.Exists(marker)) return;   // 已迁移过：尊重用户后续手动分组
                bool changed = false;
                foreach (var m in ShapeLibraryService.LoadAllItems())
                {
                    string want = null;
                    string cat = m.GetCategory();
                    if (cat == "内置素材") want = "素材";
                    else if (cat == "学术模板") want = "模板";
                    if (want != null && m.GetGroup() != want)
                    {
                        m.Group = want;
                        ShapeLibraryService.SaveMeta(m);
                        changed = true;
                    }
                }
                if (changed) ShapeLibraryService.InvalidateCachePublic();
                try { System.IO.File.WriteAllText(marker, DateTime.Now.ToString()); } catch { }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EnsureGroupAssignments");
            }
        }

        private static string AcademicMarker
        {
            get { return System.IO.Path.Combine(ShapeLibraryService.RootDir, ".academic_installed"); }
        }

        /// <summary>
        /// 安装"学术模板"成套素材（答辩封面/开题封面/研究要点页/结论页等）。
        /// 每套为单页组合收藏（分类=学术模板），供学生插入后替换文字。
        /// 独立标记，不会覆盖用户后续改动。
        /// </summary>
        public static void InstallAcademicTemplates()
        {
            try
            {
                if (System.IO.File.Exists(AcademicMarker)) return;

                var specs = new System.Collections.Generic.List<Tuple<string, string>>
                {
                    Tuple.Create("答辩封面", "defenseCover"),
                    Tuple.Create("开题封面", "proposalCover"),
                    Tuple.Create("研究要点页", "keyPoints"),
                    Tuple.Create("结论与展望", "conclusion"),
                };

                foreach (var spec in specs)
                {
                    try
                    {
                        PowerPoint.Presentation pres = null;
                        try
                        {
                            pres = App.Presentations.Add(MsoTriState.msoFalse);
                            pres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                            var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            DrawAcademicShape(slide, spec.Item2);
                            if (slide.Shapes.Count == 0) continue;

                            string id = ShapeLibraryService.NewId();
                            ShapeLibraryService.CreateItemDir(id);
                            try
                            {
                                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                                slide.Export(previewPath, "PNG", 320, 240);
                                ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                                TryDelete(previewPath);
                            }
                            catch { }

                            pres.SaveAs(ShapeLibraryService.ShapeFile(id), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                            var meta = new LibraryMeta
                            {
                                Id = id,
                                Name = "学术·" + spec.Item1,
                                Kind = "Shape",
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                ShapeCount = slide.Shapes.Count,
                                Category = "学术模板",
                                LastUsedAt = "",
                                SortOrder = Environment.TickCount
                            };
                            ShapeLibraryService.SaveMeta(meta);
                        }
                        finally
                        {
                            if (pres != null) { try { pres.Close(); } catch { } }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Academic:" + spec.Item1);
                    }
                }
                System.IO.File.WriteAllText(AcademicMarker, DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Log.Error(ex, "InstallAcademicTemplates");
            }
        }

        /// <summary>绘制学术模板（简单版式：几何块+文字，学生插入后可改字改色）。</summary>
        private static void DrawAcademicShape(PowerPoint.Slide slide, string type)
        {
            try
            {
                switch (type)
                {
                    case "defenseCover":
                        {
                            // 顶部色带 + 中部标题占位 + 底部信息条
                            var band = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 960, 24);
                            band.Fill.ForeColor.RGB = 0x1F6FA2; band.Line.Visible = MsoTriState.msoFalse;
                            var t1 = TextBoxOn(slide, "论文题目：____", 140, 220, 680, 70, 32, true);
                            var t2 = TextBoxOn(slide, "答辩人：____    导师：____    日期：____", 220, 420, 520, 40, 18, false);
                            var line = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 240, 400, 480, 2);
                            line.Fill.ForeColor.RGB = 0x1F6FA2; line.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "proposalCover":
                        {
                            var band = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 24, 540);
                            band.Fill.ForeColor.RGB = 0x2E9B4F; band.Line.Visible = MsoTriState.msoFalse;
                            var t1 = TextBoxOn(slide, "开题报告", 120, 130, 500, 90, 44, true);
                            var t2 = TextBoxOn(slide, "题目：____", 120, 260, 640, 44, 24, false);
                            var t3 = TextBoxOn(slide, "姓名/学号：____\n导师：____\n日期：____", 120, 360, 560, 110, 16, false);
                            break;
                        }
                    case "keyPoints":
                        {
                            var t = TextBoxOn(slide, "研究要点", 60, 40, 300, 50, 26, true);
                            float[] ys = { 130, 230, 330 };
                            foreach (var y in ys)
                            {
                                var c = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 80, y, 780, 70);
                                c.Fill.ForeColor.RGB = 0xE3F0FB; c.Fill.Transparency = 0;
                                c.Line.Visible = MsoTriState.msoFalse;
                                TextBoxOn(slide, "● 要点____", 110, y + 18, 700, 34, 18, false);
                            }
                            break;
                        }
                    case "conclusion":
                        {
                            TextBoxOn(slide, "结论与展望", 60, 40, 300, 50, 26, true);
                            var box = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 80, 120, 800, 260);
                            box.Fill.Visible = MsoTriState.msoFalse;
                            box.Line.ForeColor.RGB = 0x1F6FA2; box.Line.Weight = 1.5f;
                            TextBoxOn(slide, "研究结论：____\n\n不足与展望：____", 120, 160, 700, 180, 18, false);
                            TextBoxOn(slide, "谢谢聆听 · 请各位老师批评指正", 250, 460, 460, 40, 20, true);
                            break;
                        }
                }
            }
            catch { }
        }

        private static PowerPoint.Shape TextBoxOn(PowerPoint.Slide slide, string text, float left, float top, float w, float h, float size, bool bold)
        {
            var tb = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, w, h);
            var tr = tb.TextFrame.TextRange;
            tr.Text = text;
            tr.Font.Size = size;
            tr.Font.Name = "微软雅黑";
            tr.Font.Color.RGB = 0x333333;
            if (bold) tr.Font.Bold = MsoTriState.msoTrue;
            return tb;
        }

        /// <summary>在页面上绘制一个预设形状组合。</summary>
        private static void DrawBuiltinShape(PowerPoint.Slide slide, string type)
        {
            try
            {
                switch (type)
                {
                    case "card":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 20, 20, 180, 100);
                            r.Fill.ForeColor.RGB = 0x6FA21F; // BGR
                            r.Fill.Transparency = 0;
                            r.Line.Visible = MsoTriState.msoFalse;
                            r.TextFrame2.TextRange.Text = "标题";
                            r.TextFrame2.TextRange.Font.Size = 14;
                            r.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            r.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            r.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            var t = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 130, 180, 2);
                            t.Fill.ForeColor.RGB = 0x888888;
                            break;
                        }
                    case "gradientCard":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 20, 20, 180, 100);
                            r.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
                            r.Fill.GradientStops[1].Color.RGB = 0x8A2BE2;
                            r.Fill.GradientStops[2].Color.RGB = 0x1F6FA2;
                            r.Line.Visible = MsoTriState.msoFalse;
                            r.TextFrame2.TextRange.Text = "渐变";
                            r.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            r.TextFrame2.TextRange.Font.Size = 14;
                            r.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            r.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            break;
                        }
                    case "arrow":
                        {
                            var a = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, 20, 40, 180, 60);
                            a.Fill.ForeColor.RGB = 0x3C8A1F;
                            a.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "process":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 20, 180, 70);
                            r.Fill.ForeColor.RGB = 0xA26B1F;
                            r.Line.Visible = MsoTriState.msoFalse;
                            r.TextFrame2.TextRange.Text = "步骤";
                            r.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            r.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            r.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            var a = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, 210, 32, 40, 46);
                            a.Fill.ForeColor.RGB = 0x888888;
                            break;
                        }
                    case "decision":
                        {
                            var d = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeDiamond, 50, 20, 120, 100);
                            d.Fill.ForeColor.RGB = 0xA21F5C;
                            d.Line.Visible = MsoTriState.msoFalse;
                            d.TextFrame2.TextRange.Text = "条件";
                            d.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            d.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            d.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            break;
                        }
                    case "titleSub":
                        {
                            var t1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 20, 200, 50);
                            t1.Fill.Visible = MsoTriState.msoFalse;
                            t1.Line.Visible = MsoTriState.msoFalse;
                            t1.TextFrame2.TextRange.Text = "大标题";
                            t1.TextFrame2.TextRange.Font.Size = 24;
                            t1.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
                            t1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x1F6FA2;
                            var t2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 76, 200, 34);
                            t2.Fill.Visible = MsoTriState.msoFalse;
                            t2.Line.Visible = MsoTriState.msoFalse;
                            t2.TextFrame2.TextRange.Text = "副标题说明文字";
                            t2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x666666;
                            break;
                        }
                    case "iconCircle":
                        {
                            var c = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 50, 20, 100, 100);
                            c.Fill.ForeColor.RGB = 0x1F6FA2;
                            c.Line.Visible = MsoTriState.msoFalse;
                            c.TextFrame2.TextRange.Text = "A";
                            c.TextFrame2.TextRange.Font.Size = 30;
                            c.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
                            c.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            c.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            c.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            break;
                        }
                    case "badge":
                        {
                            var b = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 20, 40, 110, 34);
                            b.Fill.ForeColor.RGB = 0xA21F6F;
                            b.Line.Visible = MsoTriState.msoFalse;
                            b.TextFrame2.TextRange.Text = "新标签";
                            b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            b.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            b.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            break;
                        }
                    case "divider":
                        {
                            var d = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 54, 200, 4);
                            d.Fill.ForeColor.RGB = 0x1F6FA2;
                            d.Line.Visible = MsoTriState.msoFalse;
                            var c = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 105, 50, 12, 12);
                            c.Fill.ForeColor.RGB = 0x1F6FA2;
                            c.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "quote":
                        {
                            var v = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 20, 6, 90);
                            v.Fill.ForeColor.RGB = 0x1F6FA2;
                            v.Line.Visible = MsoTriState.msoFalse;
                            var t = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 36, 20, 180, 90);
                            t.Fill.Visible = MsoTriState.msoFalse;
                            t.Line.Visible = MsoTriState.msoFalse;
                            t.TextFrame2.TextRange.Text = "引用文字内容";
                            t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x444444;
                            t.TextFrame2.TextRange.Font.Size = 13;
                            t.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            break;
                        }
                    case "stepNumber":
                        {
                            var c = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 30, 30, 56, 56);
                            c.Fill.ForeColor.RGB = 0x1F6FA2;
                            c.Line.Visible = MsoTriState.msoFalse;
                            c.TextFrame2.TextRange.Text = "1";
                            c.TextFrame2.TextRange.Font.Size = 22;
                            c.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
                            c.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                            c.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            c.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                            var t = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 96, 42, 120, 32);
                            t.Fill.Visible = MsoTriState.msoFalse;
                            t.Line.Visible = MsoTriState.msoFalse;
                            t.TextFrame2.TextRange.Text = "步骤说明";
                            t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x444444;
                            t.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            break;
                        }
                    case "metric":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 20, 140, 100);
                            r.Fill.ForeColor.RGB = 0xEEF4FA;
                            r.Line.ForeColor.RGB = 0x1F6FA2;
                            r.Line.Weight = 1;
                            var n = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 30, 30, 120, 46);
                            n.Fill.Visible = MsoTriState.msoFalse;
                            n.Line.Visible = MsoTriState.msoFalse;
                            n.TextFrame2.TextRange.Text = "88%";
                            n.TextFrame2.TextRange.Font.Size = 28;
                            n.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
                            n.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x1F6FA2;
                            var l = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 30, 80, 120, 26);
                            l.Fill.Visible = MsoTriState.msoFalse;
                            l.Line.Visible = MsoTriState.msoFalse;
                            l.TextFrame2.TextRange.Text = "完成率";
                            l.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x888888;
                            break;
                        }
                    case "ring":
                        {
                            var o1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 50, 20, 100, 100);
                            o1.Fill.Visible = MsoTriState.msoFalse;
                            o1.Line.ForeColor.RGB = 0x1F6FA2;
                            o1.Line.Weight = 8;
                            var o2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 72, 42, 56, 56);
                            o2.Fill.Visible = MsoTriState.msoFalse;
                            o2.Line.ForeColor.RGB = 0xA21F6F;
                            o2.Line.Weight = 4;
                            break;
                        }
                    case "bulletList":
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                var c = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 20, 20 + i * 34, 12, 12);
                                c.Fill.ForeColor.RGB = 0x1F6FA2;
                                c.Line.Visible = MsoTriState.msoFalse;
                                var t = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 40, 16 + i * 34, 180, 20);
                                t.Fill.Visible = MsoTriState.msoFalse;
                                t.Line.Visible = MsoTriState.msoFalse;
                                t.TextFrame2.TextRange.Text = "要点 " + (i + 1);
                                t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x444444;
                            }
                            break;
                        }
                    case "sidebar":
                        {
                            var s = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 20, 20, 34, 140);
                            s.Fill.ForeColor.RGB = 0x1F6FA2;
                            s.Line.Visible = MsoTriState.msoFalse;
                            var t = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 64, 20, 160, 140);
                            t.Fill.Visible = MsoTriState.msoFalse;
                            t.Line.Visible = MsoTriState.msoFalse;
                            t.TextFrame2.TextRange.Text = "侧边栏内容";
                            t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x444444;
                            t.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            break;
                        }
                }
            }
            catch { }
        }

        // ---------------- 导出为 PPTX（C2） ----------------

        /// <summary>把所有形状收藏导出为 .pptx（每页一个收藏）。</summary>
        public static bool ExportLibraryToPptx(string path)
        {
            PowerPoint.Presentation outPres = null;
            System.Windows.Forms.IDataObject clip = ClipboardGuard.Snapshot();   // 保护剪贴板
            try
            {
                outPres = App.Presentations.Add(MsoTriState.msoFalse);
                outPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                int added = 0;

                foreach (var meta in ShapeLibraryService.LoadAllItems())
                {
                    if (meta.Kind != "Shape") continue;
                    string shapePath = ShapeLibraryService.ShapeFile(meta.Id);
                    if (!System.IO.File.Exists(shapePath)) continue;

                    PowerPoint.Presentation libPres = null;
                    try
                    {
                        libPres = App.Presentations.Open(shapePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                        var src = libPres.Slides[1];
                        if (src.Shapes.Count == 0) continue;

                        var idx = new System.Collections.Generic.List<int>();
                        for (int i = 1; i <= src.Shapes.Count; i++)
                        {
                            if (!IsPlaceholder(src.Shapes[i])) idx.Add(i);
                        }
                        if (idx.Count == 0) continue;

                        var range = src.Shapes.Range(idx.ToArray());

                        var page = outPres.Slides.Add(outPres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                        var pasted = CopyPasteWithRetry(() => range.Copy(), page.Shapes);
                        float minL = float.MaxValue, minT = float.MaxValue;
                        foreach (PowerPoint.Shape s in pasted)
                        {
                            minL = Math.Min(minL, (float)s.Left);
                            minT = Math.Min(minT, (float)s.Top);
                        }
                        foreach (PowerPoint.Shape s in pasted)
                        {
                            s.Left -= minL;
                            s.Top -= minT;
                        }
                        // 页面备注写名称
                        try
                        {
                            page.NotesPage.Shapes.Placeholders[(int)PowerPoint.PpPlaceholderType.ppPlaceholderBody]
                                .TextFrame.TextRange.Text = meta.Name;
                        }
                        catch { }
                        added++;
                    }
                    finally
                    {
                        if (libPres != null) { try { libPres.Close(); } catch { } }
                    }
                }

                if (added == 0)
                {
                    System.Windows.Forms.MessageBox.Show("没有可导出的形状收藏。", "素材库");
                    return false;
                }
                outPres.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("导出失败：" + ex.Message, "素材库");
                return false;
            }
            finally
            {
                ClipboardGuard.Restore(clip);
                if (outPres != null) { try { outPres.Close(); } catch { } }
            }
        }

        // ---------------- 快速调整（插入后微调条） ----------------

        /// <summary>对当前选中形状执行微调操作。</summary>
        public static void AdjustSelected(string action)
        {
            try
            {
                var sel = App.ActiveWindow.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.ShapeRange == null || sel.ShapeRange.Count == 0) return;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    switch (action)
                    {
                        case "rotateL": s.Rotation -= 15; break;
                        case "rotateR": s.Rotation += 15; break;
                        case "opacityUp":
                            if (s.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                                s.Fill.Transparency = Math.Max(0, s.Fill.Transparency - 0.1f);
                            break;
                        case "opacityDown":
                            if (s.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                                s.Fill.Transparency = Math.Min(1, s.Fill.Transparency + 0.1f);
                            break;
                        case "nudgeL": s.Left -= 5; break;
                        case "nudgeR": s.Left += 5; break;
                        case "nudgeU": s.Top -= 5; break;
                        case "nudgeD": s.Top += 5; break;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 撤销最近一次文档级操作（插入/应用样式/调整条/格式化等）。
        /// 经 CommandBars.ExecuteMso("Undo") 触发 PowerPoint 原生撤销栈。
        /// </summary>
        public static void UndoLastAction()
        {
            try
            {
                App.CommandBars.ExecuteMso("Undo");
            }
            catch { /* 无可撤销操作或处于不可撤销上下文时静默 */ }
        }

        /// <summary>重做（撤销的反向）。ExecuteMso("Redo")。</summary>
        public static void RedoLastAction()
        {
            try
            {
                App.CommandBars.ExecuteMso("Redo");
            }
            catch { }
        }

        /// <summary>
        /// 复制→粘贴（带重试）：PowerPoint 的 Copy 是异步的，立即 Paste 可能报
        /// "Clipboard is empty"；逐次延迟并重新 Copy，最多 5 次。
        /// </summary>
        private static PowerPoint.ShapeRange CopyPasteWithRetry(Action copy, PowerPoint.Shapes target)
        {
            Exception last = null;
            for (int attempt = 1; attempt <= 5; attempt++)
            {
                try
                {
                    copy();
                    System.Threading.Thread.Sleep(40 * attempt);
                    // Copy 后校验剪贴板确已写入 PowerPoint 形状数据；
                    // 若为空（Copy 静默失败/被其他程序抢占）继续重试，避免 Paste 报 "Clipboard is empty"
                    bool hasData = false;
                    try
                    {
                        hasData = System.Windows.Forms.Clipboard.ContainsData("PowerPoint Shapedata")
                            || System.Windows.Forms.Clipboard.ContainsData("PNG");
                    }
                    catch { }
                    if (!hasData && attempt < 5)
                    {
                        Log.Write("copy check: clipboard lacks PPT data (attempt " + attempt + "), retry");
                        continue;
                    }
                    return target.Paste();
                }
                catch (Exception ex)
                {
                    last = ex;
                    Log.Write("paste retry " + attempt + ": " + ex.Message);
                }
            }
            throw last ?? new Exception("粘贴失败");
        }

        // ---------------- 工具 ----------------

        private static bool IsPlaceholder(PowerPoint.Shape shp)
        {
            try
            {
                return shp.PlaceholderFormat != null;
            }
            catch { return false; }
        }

        /// <summary>取形状主色（"RRGGBB"）：优先纯色填充，其次渐变首色，再次线条色；取不到返回 null。</summary>
        private static string GetPrimaryColorHex(PowerPoint.ShapeRange range)
        {
            try
            {
                if (range == null || range.Count == 0) return null;
                var shp = range[1];
                try
                {
                    int ft = (int)shp.Fill.Type;
                    if (ft == 1)   // msoFillSolid
                    {
                        return ShapeStyleExtractor.RgbToHex(shp.Fill.ForeColor.RGB);
                    }
                    if (ft == 3)   // msoFillGradient
                    {
                        var stops = shp.Fill.GradientStops;
                        if (stops != null && stops.Count > 0)
                            return ShapeStyleExtractor.RgbToHex(stops[1].Color.RGB);
                    }
                }
                catch { }
                try
                {
                    if (shp.Line.Visible == MsoTriState.msoTrue)
                        return ShapeStyleExtractor.RgbToHex(shp.Line.ForeColor.RGB);
                }
                catch { }
            }
            catch { }
            return null;
        }

        /// <summary>占位符且完全无内容（无文字/无表格/无图表/非图片型）——即"单击此处添加…"空引导框。
        /// 收藏整页/整稿时应跳过这类空占位，但保留含文字图片内容的占位符（那才是页面上真正的内容）。</summary>
        private static bool IsEmptyPlaceholder(PowerPoint.Shape shp)
        {
            try
            {
                if (!IsPlaceholder(shp)) return false;
                // 图片/图表/表格型占位：有内容载体 → 保留
                try
                {
                    switch (shp.Type)
                    {
                        case MsoShapeType.msoPicture:
                        case MsoShapeType.msoLinkedPicture:
                        case MsoShapeType.msoChart:
                        case MsoShapeType.msoTable:
                            return false;
                    }
                }
                catch { }
                // 文字型占位：有文字则保留
                try
                {
                    var tf = shp.TextFrame2;
                    if (tf != null && tf.HasText == MsoTriState.msoTrue && tf.TextRange != null && tf.TextRange.Length > 0)
                        return false;
                }
                catch { }
                return true;   // 空占位引导框
            }
            catch { return true; }
        }

        private static int GetAutoType(PowerPoint.Shape shape)
        {
            try { return (int)shape.AutoShapeType; }
            catch { return -1; }
        }

        /// <summary>根据形状类型自动生成中文名。</summary>
        private static string GetAutoName(PowerPoint.ShapeRange range)
        {
            try
            {
                if (range.Count == 1)
                {
                    switch (range[1].AutoShapeType)
                    {
                        case MsoAutoShapeType.msoShapeRectangle: return "矩形";
                        case MsoAutoShapeType.msoShapeRoundedRectangle: return "圆角矩形";
                        case MsoAutoShapeType.msoShapeOval: return "椭圆";
                        case MsoAutoShapeType.msoShapeRightTriangle: return "直角三角形";
                        case MsoAutoShapeType.msoShapeDiamond: return "菱形";
                        case MsoAutoShapeType.msoShapeParallelogram: return "平行四边形";
                        case MsoAutoShapeType.msoShapeHexagon: return "六边形";
                        case MsoAutoShapeType.msoShapePentagon: return "五边形";
                        case MsoAutoShapeType.msoShapeRightArrow: return "右箭头";
                        case MsoAutoShapeType.msoShapeLeftArrow: return "左箭头";
                        case MsoAutoShapeType.msoShapeUpArrow: return "上箭头";
                        case MsoAutoShapeType.msoShapeDownArrow: return "下箭头";
                        case MsoAutoShapeType.msoShapeLeftRightArrow: return "左右箭头";
                        case MsoAutoShapeType.msoShapeUpDownArrow: return "上下箭头";
                        case MsoAutoShapeType.msoShapeChevron: return "燕尾形";
                        case MsoAutoShapeType.msoShapeRegularPentagon: return "五边形";
                        case MsoAutoShapeType.msoShapeCircularArrow: return "环形箭头";
                        case MsoAutoShapeType.msoShapeNotPrimitive: return "自由形状";
                        case MsoAutoShapeType.msoShapeRoundedRectangularCallout: return "圆角标注";
                        case MsoAutoShapeType.msoShapeLineCallout1: return "标注";
                        case MsoAutoShapeType.msoShapeFlowchartProcess: return "流程图框";
                        case MsoAutoShapeType.msoShapeFlowchartDecision: return "流程图判定";
                        case MsoAutoShapeType.msoShapeCloud: return "云朵";
                        case MsoAutoShapeType.msoShapeHeart: return "心形";
                        case MsoAutoShapeType.msoShape5pointStar: return "五角星";
                        case MsoAutoShapeType.msoShape8pointStar: return "八角星";
                        case MsoAutoShapeType.msoShapeMoon: return "月牙";
                        case MsoAutoShapeType.msoShapeLightningBolt: return "闪电";
                        case MsoAutoShapeType.msoShapeSun: return "太阳";
                        case MsoAutoShapeType.msoShapeTrapezoid: return "梯形";
                        case MsoAutoShapeType.msoShapeDonut: return "环形";
                        case MsoAutoShapeType.msoShapePie: return "扇形";
                        case MsoAutoShapeType.msoShapeArc: return "弧形";
                        case MsoAutoShapeType.msoShapeBlockArc: return "拱形";
                        default:
                            return "形状";
                    }
                }
                if (range.Count > 1) return "组合(" + range.Count + ")";
            }
            catch { }
            return "形状";
        }

        private static string GetSourceInfo()
        {
            try
            {
                if (App.ActivePresentation == null) return "";
                return App.ActivePresentation.FullName + " · 幻灯片 " + App.ActiveWindow.View.Slide.SlideIndex;
            }
            catch { return ""; }
        }

        private static void TryDelete(string path)
        {
            try { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); } catch { }
        }

        // ---------------- 预览重建（修复历史错位缩略图） ----------------

        /// <summary>重建单个条目的预览缩略图：打开其库文稿，从归一化形状自然尺寸导出（与插入一致）。</summary>
        public static bool RebuildPreview(string id)
        {
            if (!ShapeLibraryService.IsSafeId(id)) return false;
            string shapePath = ShapeLibraryService.ShapeFile(id);
            if (!System.IO.File.Exists(shapePath)) return false;
            PowerPoint.Presentation libPres = null;
            try
            {
                libPres = App.Presentations.Open(shapePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                var slide = libPres.Slides[1];
                if (slide.Shapes.Count == 0) return false;
                var idx = new System.Collections.Generic.List<int>();
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    if (!IsPlaceholder(slide.Shapes[i])) idx.Add(i);
                }
                if (idx.Count == 0) return false;
                var range = slide.Shapes.Range(idx.ToArray());
                string tmp = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_rebuild_tmp.png");
                range.Export(tmp, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                if (System.IO.File.Exists(tmp))
                {
                    ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(tmp));
                    TryDelete(tmp);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "RebuildPreview " + id);
                return false;
            }
            finally
            {
                if (libPres != null) { try { libPres.Close(); } catch { } }
            }
        }

        /// <summary>重建全部形状条目的预览缩略图。返回成功数。</summary>
        public static int RebuildAllPreviews()
        {
            int ok = 0;
            foreach (var m in ShapeLibraryService.LoadAllItems())
            {
                if (m.Kind != "Shape") continue;
                try { if (RebuildPreview(m.Id)) ok++; }
                catch { }
            }
            return ok;
        }


        // ---------------- 图片素材：导入入库 + 程序生成示意 ---------------

        /// <summary>把本地图片文件入库为 素材>图片 收藏（单图单页 pptx）。返回条目 id。</summary>
        public static string ImportPictureToLibrary(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath)) return null;
            PowerPoint.Presentation tempPres = null;
            try
            {
                // 先读原图像素尺寸，按等比（长边≤640）计算目标尺寸，避免 AddPicture 固定尺寸拉伸变形
                int srcW = 640, srcH = 360;
                try
                {
                    using (var img = System.Drawing.Image.FromFile(filePath))
                    {
                        srcW = img.Width; srcH = img.Height;
                    }
                }
                catch { }
                if (srcW <= 0 || srcH <= 0) { srcW = 640; srcH = 360; }
                float scale = System.Math.Min(1f, System.Math.Min(640f / srcW, 360f / srcH));
                int dw = System.Math.Max(1, (int)(srcW * scale));
                int dh = System.Math.Max(1, (int)(srcH * scale));

                tempPres = App.Presentations.Add(MsoTriState.msoFalse);
                tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                var slide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                var pic = slide.Shapes.AddPicture(filePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, dw, dh);
                if (pic == null) return null;

                string id = ShapeLibraryService.NewId();
                ShapeLibraryService.CreateItemDir(id);
                try
                {
                    string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                    slide.Export(previewPath, "PNG", 320, 240);
                    ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                    TryDelete(previewPath);
                }
                catch { }
                tempPres.SaveAs(ShapeLibraryService.ShapeFile(id), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                var meta = new LibraryMeta
                {
                    Id = id,
                    Name = System.IO.Path.GetFileNameWithoutExtension(filePath),
                    Kind = "Shape",
                    CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ShapeCount = 1,
                    Category = "图片",
                    Group = "素材",
                    LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SortOrder = Environment.TickCount
                };
                ShapeLibraryService.SaveMeta(meta);
                return id;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ImportPicture");
                return null;
            }
            finally
            {
                if (tempPres != null) { try { tempPres.Close(); } catch { } }
            }
        }

        /// <summary>程序生成示意图片（PNG 字节→临时文件→入库 素材>图片）。离线、无版权。独立 marker 幂等。</summary>
        public static void GenerateSamplePictures()
        {
            try
            {
                string marker = System.IO.Path.Combine(ShapeLibraryService.RootDir, ".samples_installed");
                if (System.IO.File.Exists(marker)) return;
                var types = new System.Collections.Generic.List<Tuple<string, string>>  // name, kind
                {
                    Tuple.Create("示意·蓝紫渐变", "grad"),
                    Tuple.Create("示意·斜纹纹理", "stripes"),
                    Tuple.Create("示意·圆点纹理", "dots"),
                    Tuple.Create("示意·网格底图", "grid"),
                };
                foreach (var t in types)
                {
                    try
                    {
                        string tmp = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, ShapeLibraryService.NewId() + "_gen.png");
                        DrawSamplePng(tmp, t.Item2);
                        if (!System.IO.File.Exists(tmp)) continue;
                        string id = ImportPictureToLibrary(tmp);
                        TryDelete(tmp);
                        if (id != null)
                        {
                            // 改名带“示意·”
                            var m = ShapeLibraryService.LoadMeta(id);
                            if (m != null) { m.Name = t.Item1; ShapeLibraryService.SaveMeta(m); }
                        }
                    }
                    catch (Exception ex) { Log.Error(ex, "Sample:" + t.Item1); }
                }
                try
                {
                    System.IO.File.WriteAllText(
                        System.IO.Path.Combine(ShapeLibraryService.RootDir, ".samples_installed"),
                        DateTime.Now.ToString());
                }
                catch { }
            }
            catch (Exception ex) { Log.Error(ex, "GenerateSamplePictures"); }
        }

        private static void DrawSamplePng(string path, string kind)
        {
            try
            {
                int w = 640, h = 360;
                using (var bmp = new System.Drawing.Bitmap(w, h))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        switch (kind)
                        {
                            case "grad":
                                {
                                    var r = new System.Drawing.Rectangle(0, 0, w, h);
                                    using (var br = new System.Drawing.Drawing2D.LinearGradientBrush(r,
                                        System.Drawing.Color.FromArgb(0x8A2BE2), System.Drawing.Color.FromArgb(0x1F6FA2), 60f))
                                    { g.FillRectangle(br, r); }
                                    break;
                                }
                            case "stripes":
                                {
                                    g.Clear(System.Drawing.Color.White);
                                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(0x1F6FA2), 10))
                                    {
                                        for (int x = -h; x < w; x += 46)
                                            g.DrawLine(pen, x, h, x + h, 0);
                                    }
                                    break;
                                }
                            case "dots":
                                {
                                    g.Clear(System.Drawing.Color.FromArgb(0xEEF2F7));
                                    using (var br = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(0x9FB6CC)))
                                    {
                                        for (int y = 20; y < h; y += 44)
                                            for (int x = 20; x < w; x += 44)
                                                g.FillEllipse(br, x, y, 14, 14);
                                    }
                                    break;
                                }
                            case "grid":
                                {
                                    g.Clear(System.Drawing.Color.FromArgb(0xF7F9FB));
                                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(0xD5DEE8), 1))
                                    {
                                        for (int x = 0; x <= w; x += 40) g.DrawLine(pen, x, 0, x, h);
                                        for (int y = 0; y <= h; y += 40) g.DrawLine(pen, 0, y, w, y);
                                    }
                                    using (var br = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(0x1F6FA2)))
                                    {
                                        g.FillRectangle(br, 0, 0, 12, h);
                                    }
                                    break;
                                }
                        }
                    }
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            catch (Exception ex) { Log.Error(ex, "DrawSamplePng"); }
        }

        // ---------------- 内置内容库（背景/文字/图例 等） ----------------

        private static string ContentMarker
        {
            get { return System.IO.Path.Combine(ShapeLibraryService.RootDir, ".content_installed"); }
        }

        /// <summary>安装"内容库"条目：按一级组 + 二级分类归位（背景/文字/图例；素材补充边框线条）。幂等。</summary>
        public static void InstallLibraryContent()
        {
            try
            {
                if (System.IO.File.Exists(ContentMarker)) return;

                var specs = new System.Collections.Generic.List<Tuple<string, string, string, string>> // name, group, category, kind
                {
                    // 背景
                    Tuple.Create("背景·学术蓝", "背景", "纯色", "bg"),
                    Tuple.Create("背景·奶油白", "背景", "纯色", "bg"),
                    Tuple.Create("背景·蓝紫渐变", "背景", "渐变", "bgGrad"),
                    Tuple.Create("背景·青绿渐变", "背景", "渐变", "bgGrad2"),
                    // 文字
                    Tuple.Create("文字·大标题", "文字", "标题", "txtTitle"),
                    Tuple.Create("文字·副标题", "文字", "副标题", "txtSub"),
                    Tuple.Create("文字·数据大数字", "文字", "特效字", "txtNumber"),
                    Tuple.Create("文字·引用块", "文字", "特效字", "txtQuote"),
                    // 图例
                    Tuple.Create("图例·流程箭头", "图例", "流程", "legendFlow"),
                    Tuple.Create("图例·环形示意", "图例", "环形", "legendRing"),
                    Tuple.Create("图例·对比双向", "图例", "对比", "legendCompare"),
                    // 素材补充（线条/边框）
                    Tuple.Create("素材·双线分隔", "素材", "线条", "matLine"),
                    Tuple.Create("素材·圆角边框", "素材", "边框", "matFrame"),
                };

                foreach (var spec in specs)
                {
                    try
                    {
                        PowerPoint.Presentation pres = null;
                        try
                        {
                            pres = App.Presentations.Add(MsoTriState.msoFalse);
                            pres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                            var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            DrawContentItem(slide, spec.Item4);
                            if (slide.Shapes.Count == 0) continue;

                            string id = ShapeLibraryService.NewId();
                            ShapeLibraryService.CreateItemDir(id);
                            try
                            {
                                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                                slide.Export(previewPath, "PNG", 320, 240);
                                ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                                TryDelete(previewPath);
                            }
                            catch { }

                            pres.SaveAs(ShapeLibraryService.ShapeFile(id), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

                            var meta = new LibraryMeta
                            {
                                Id = id,
                                Name = spec.Item1,
                                Kind = "Shape",
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                ShapeCount = slide.Shapes.Count,
                                Category = spec.Item3,
                                Group = spec.Item2,
                                LastUsedAt = "",
                                SortOrder = Environment.TickCount
                            };
                            ShapeLibraryService.SaveMeta(meta);
                        }
                        finally
                        {
                            if (pres != null) { try { pres.Close(); } catch { } }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Content:" + spec.Item1);
                    }
                }
                // 程序生成示意图片 → 素材>图片
                GenerateSamplePictures();
                System.IO.File.WriteAllText(ContentMarker, DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Log.Error(ex, "InstallLibraryContent");
            }
        }

        /// <summary>绘制内容库条目（画布 960×540 比例，入库归一化）。</summary>
        private static void DrawContentItem(PowerPoint.Slide slide, string kind)
        {
            try
            {
                switch (kind)
                {
                    case "bg":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 960, 540);
                            r.Fill.ForeColor.RGB = 0x1F6FA2; r.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "bgGrad":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 960, 540);
                            r.Fill.TwoColorGradient(MsoGradientStyle.msoGradientDiagonalDown, 1);
                            try { r.Fill.GradientStops[1].Color.RGB = 0x8A2BE2; r.Fill.GradientStops[2].Color.RGB = 0x1F6FA2; } catch { }
                            r.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "bgGrad2":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 960, 540);
                            r.Fill.TwoColorGradient(MsoGradientStyle.msoGradientDiagonalUp, 1);
                            try { r.Fill.GradientStops[1].Color.RGB = 0x2E9B4F; r.Fill.GradientStops[2].Color.RGB = 0x0B7A9B; } catch { }
                            r.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "txtTitle":
                        {
                            TextBoxOn(slide, "主 标 题", 60, 150, 840, 200, 88, true);
                            var line = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 90, 330, 180, 6);
                            line.Fill.ForeColor.RGB = 0x1F6FA2; line.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "txtSub":
                        {
                            TextBoxOn(slide, "副标题 · 说明文字", 60, 200, 700, 70, 36, true);
                            var line = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 62, 300, 420, 2);
                            line.Fill.ForeColor.RGB = 0x9AA7B4; line.Line.Visible = MsoTriState.msoFalse;
                            TextBoxOn(slide, "补充描述内容可以写在这里", 62, 330, 560, 44, 20, false);
                            break;
                        }
                    case "txtNumber":
                        {
                            var n = TextBoxOn(slide, "88%", 60, 140, 840, 220, 120, true);
                            var lbl = TextBoxOn(slide, "数据说明标签", 66, 350, 400, 40, 22, false);
                            break;
                        }
                    case "txtQuote":
                        {
                            var bar = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 60, 120, 8, 160);
                            bar.Fill.ForeColor.RGB = 0x1F6FA2; bar.Line.Visible = MsoTriState.msoFalse;
                            TextBoxOn(slide, "“引用或强调的核心观点文字”", 96, 120, 700, 100, 30, false);
                            TextBoxOn(slide, "—— 出处 / 作者", 96, 230, 400, 40, 18, false);
                            break;
                        }
                    case "legendFlow":
                        {
                            float x = 60;
                            string[] steps = { "需求", "设计", "实现", "验证" };
                            for (int i = 0; i < steps.Length; i++)
                            {
                                var b = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, x, 200, 150, 66);
                                b.Fill.ForeColor.RGB = i % 2 == 0 ? 0x1F6FA2 : 0x2E9B4F;
                                b.Line.Visible = MsoTriState.msoFalse;
                                var tr = b.TextFrame2.TextRange; tr.Text = steps[i];
                                tr.Font.Size = 18; tr.Font.Fill.ForeColor.RGB = 0xFFFFFF;
                                tr.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                                if (i < steps.Length - 1)
                                {
                                    var a = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, x + 150, 220, 44, 28);
                                    a.Fill.ForeColor.RGB = 0x666666; a.Line.Visible = MsoTriState.msoFalse;
                                }
                                x += 214;
                            }
                            break;
                        }
                    case "legendRing":
                        {
                            var o = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 200, 150, 220, 220);
                            o.Fill.Visible = MsoTriState.msoFalse;
                            o.Line.ForeColor.RGB = 0x1F6FA2; o.Line.Weight = 14;
                            var inner = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 276, 226, 68, 68);
                            inner.Fill.ForeColor.RGB = 0x2E9B4F; inner.Line.Visible = MsoTriState.msoFalse;
                            TextBoxOn(slide, "中心主题", 590, 220, 260, 80, 24, true);
                            TextBoxOn(slide, "外围说明文字区域用于补充内容", 590, 300, 260, 90, 15, false);
                            break;
                        }
                    case "legendCompare":
                        {
                            TextBoxOn(slide, "方案 A", 120, 150, 280, 60, 26, true);
                            TextBoxOn(slide, "方案 B", 560, 150, 280, 60, 26, true);
                            var a = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, 120, 260, 300, 40);
                            a.Fill.ForeColor.RGB = 0x2E9B4F; a.Line.Visible = MsoTriState.msoFalse;
                            var b = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeLeftArrow, 560, 260, 300, 40);
                            b.Fill.ForeColor.RGB = 0xD63A2A; b.Line.Visible = MsoTriState.msoFalse;
                            TextBoxOn(slide, "优势说明……", 120, 330, 300, 60, 16, false);
                            TextBoxOn(slide, "优势说明……", 560, 330, 300, 60, 16, false);
                            break;
                        }
                    case "matLine":
                        {
                            var l1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 60, 250, 840, 3);
                            l1.Fill.ForeColor.RGB = 0x2E9B4F; l1.Line.Visible = MsoTriState.msoFalse;
                            var l2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 60, 280, 840, 1);
                            l2.Fill.ForeColor.RGB = 0x9AA7B4; l2.Line.Visible = MsoTriState.msoFalse;
                            break;
                        }
                    case "matFrame":
                        {
                            var r = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 60, 150, 400, 240);
                            r.Fill.Visible = MsoTriState.msoFalse;
                            r.Line.ForeColor.RGB = 0x1F6FA2; r.Line.Weight = 2.5f;
                            TextBoxOn(slide, "框内内容占位", 150, 250, 240, 50, 18, false);
                            break;
                        }
                }
            }
            catch { }
        }
    }
}
