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
        public static string CollectWholePage()
        {
            EnsurePresentation();
            try
            {
                var pres = App.ActivePresentation;
                if (pres == null) return null;
                var slide = pres.Slides[App.ActiveWindow.View.Slide.SlideIndex];
                var idx = new System.Collections.Generic.List<int>();
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    if (!IsPlaceholder(slide.Shapes[i])) idx.Add(i);
                }
                if (idx.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("当前页没有可收藏的形状。", "PPT形状库",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return null;
                }
                var range = slide.Shapes.Range(idx.ToArray());
                return CollectRange(range, null);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("收藏整页失败：" + ex.Message, "PPT形状库");
                return null;
            }
        }

        /// <summary>收藏任意 ShapeRange 的核心实现。</summary>
        public static string CollectRange(PowerPoint.ShapeRange range, string name)
        {
            if (range == null || range.Count == 0) return null;

            PowerPoint.Presentation tempPres = null;
            string id = ShapeLibraryService.NewId();

            try
            {
                // 去重提示
                try
                {
                    string dup = ShapeLibraryService.FindDuplicate(range);
                    if (dup != null)
                    {
                        if (System.Windows.Forms.MessageBox.Show(
                            "检测到可能重复的收藏：「" + dup + "」。仍要收藏吗？",
                            "PPT形状库", System.Windows.Forms.MessageBoxButtons.YesNo,
                            System.Windows.Forms.MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
                        {
                            return null;
                        }
                    }
                }
                catch { }

                // 0. 创建收藏目录
                ShapeLibraryService.CreateItemDir(id);

                // 1. 生成缩略图：直接导出选中形状
                string previewPath = System.IO.Path.Combine(ShapeLibraryService.ItemsDir, id + "_tmp.png");
                try
                {
                    range.Export(previewPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 640, 480);
                    ShapeLibraryService.SavePreviewImage(id, System.IO.File.ReadAllBytes(previewPath));
                }
                catch { /* 缩略图失败不阻断 */ }
                finally { TryDelete(previewPath); }

                // 2. 新建隐藏演示文稿，粘贴形状
                tempPres = App.Presentations.Add(MsoTriState.msoFalse);
                tempPres.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
                var slide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                var pasted = CopyPasteWithRetry(() => range.Copy(), slide.Shapes);
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
                    Category = "未分类",
                    Tags = "",
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
                System.Windows.Forms.MessageBox.Show("收藏失败：" + ex.Message, "PPTShapeLibrary");
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
                    range.Export(previewPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 640, 480);
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
                    Category = "未分类",
                    Tags = "",
                    LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    SortOrder = Environment.TickCount,
                    SourceAutoType = GetAutoType(range[1])
                };
                ShapeLibraryService.SaveMeta(meta);
                return id;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("收藏样式失败：" + ex.Message, "PPTShapeLibrary");
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
            try
            {
                EnsurePresentation();
            }
            catch (InvalidOperationException ex)
            {
                error = ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }
            string path = ShapeLibraryService.ShapeFile(id);
            if (!System.IO.File.Exists(path))
            {
                error = "收藏文件不存在，可能已被移动或删除。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }

            PowerPoint.Presentation libPres = null;
            try
            {
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
                var pasted = CopyPasteWithRetry(() => range.Copy(), targetShapes);

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
                    try
                    {
                        pasted.ScaleWidth(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                        pasted.ScaleHeight(scale, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);
                    }
                    catch { }
                    // 缩放后重新计算边界
                    minL = float.MaxValue; minT = float.MaxValue; maxR = float.MinValue; maxB = float.MinValue;
                    foreach (PowerPoint.Shape s in pasted)
                    {
                        minL = Math.Min(minL, (float)s.Left);
                        minT = Math.Min(minT, (float)s.Top);
                        maxR = Math.Max(maxR, (float)(s.Left + s.Width));
                        maxB = Math.Max(maxB, (float)(s.Top + s.Height));
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
                return true;
            }
            catch (Exception ex)
            {
                error = "插入失败：" + ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }
            finally
            {
                if (libPres != null)
                {
                    try { libPres.Close(); } catch { }
                }
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
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }
            string json = ShapeLibraryService.LoadStyleJson(id);
            if (json == null)
            {
                error = "样式文件不存在。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }
            var selection = App.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange == null || selection.ShapeRange.Count == 0)
            {
                error = "请先在页面上选中要应用样式的形状。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
                return false;
            }
            var preset = ShapeStyleExtractor.FromJson(json);
            if (preset == null)
            {
                error = "样式解析失败。";
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
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
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPTShapeLibrary");
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
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPT形状库");
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
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPT形状库");
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
                    if (showError) System.Windows.Forms.MessageBox.Show(error, "PPT形状库");
                    return false;
                }
                if (showError)
                {
                    System.Windows.Forms.MessageBox.Show("已应用到 " + applied + " 个同类型形状。", "PPT形状库");
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "批量应用失败：" + ex.Message;
                if (showError) System.Windows.Forms.MessageBox.Show(error, "PPT形状库");
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
                                Tags = "",
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
                    System.Windows.Forms.MessageBox.Show("没有可导出的形状收藏。", "PPT形状库");
                    return false;
                }
                outPres.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("导出失败：" + ex.Message, "PPT形状库");
                return false;
            }
            finally
            {
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
                    System.Threading.Thread.Sleep(90 * attempt);
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
    }
}
