using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PPTShapeLibrary
{
    /// <summary>WPF 亚克力风格收藏面板。</summary>
    public partial class ShapeLibraryPaneWpf : UserControl, ILibraryPane
    {
        private readonly Dictionary<string, Border> _cards = new Dictionary<string, Border>();
        private readonly HashSet<string> _selectedIds = new HashSet<string>();
        private static readonly Dictionary<string, BitmapSource> _previewCache = new Dictionary<string, BitmapSource>();

        private string _filter = "";
        private string _categoryFilter = "全部";
        private string _sortMode = "最新";
        private bool _updatingUI = false;
        private int _loadedCount = 200;
        private int _pageSize = 200;
        private bool _isDark = false;
        private Settings _settings = new Settings();

        private static readonly string[] IconText = { "\uE8A5" };

        public ShapeLibraryPaneWpf()
        {
            InitializeComponent();
            try
            {
                _settings = ShapeLibraryService.LoadSettings();
                _pageSize = _settings.PageSize > 0 ? _settings.PageSize : 200;
                _loadedCount = _pageSize;
                cmbSort.Items.Add("最新");
                cmbSort.Items.Add("最旧");
                cmbSort.Items.Add("名称");
                cmbSort.Items.Add("最近使用");
                cmbSort.Items.Add("最常用");
                cmbSort.Items.Add("自定义");
                cmbSort.SelectedIndex = 0;
                ApplyTheme();
                RefreshLibrary();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WpfPane ctor");
                throw;
            }
        }

        public void ApplySettings(Settings s)
        {
            _settings = s;
            _pageSize = s.PageSize > 0 ? s.PageSize : 200;
            _loadedCount = _pageSize;
            ApplyTheme();
            RefreshLibrary();
        }

        // ---------------- 主题 ----------------

        private void ApplyTheme()
        {
            _isDark = _settings.Theme == "dark";
            SetColor("BgBrush", _isDark ? "#171A1E" : "#F2F4F7");
            SetColor("CardBrush", _isDark ? "#23272E" : "#FFFFFF");
            SetColor("TextBrush", _isDark ? "#E6EAF0" : "#1B1F23");
            SetColor("SubTextBrush", _isDark ? "#8A94A3" : "#7A8494");
            SetColor("BorderBrush", _isDark ? "#333A44" : "#E3E8EF");
            SetColor("ChipBg", _isDark ? "#2C323B" : "#E9EDF3");
            SetColor("ChipHover", _isDark ? "#39414D" : "#DCE4EE");
            SetColor("AccentBrush", _isDark ? "#4DA3FF" : "#0067C0");
            SetColor("AccentHover", _isDark ? "#3D8FE0" : "#005AA8");
            SetColor("AccentSoft", _isDark ? "#1E2A38" : "#E8F2FB");
            SetColor("PurpleBrush", "#B06FD0");
        }

        private void SetColor(string key, string hex)
        {
            try
            {
                var brush = (SolidColorBrush)FindResource(key);
                brush.Color = (Color)ColorConverter.ConvertFromString(hex);
            }
            catch { }
        }

        // ---------------- 数据 ----------------

        public void RefreshLibrary()
        {
            try
            {
                RefreshCore();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WpfPane Refresh");
            }
        }

        private void RefreshCore()
        {
            var all = ShapeLibraryService.LoadAllItems();

            // 分类标签
            _updatingUI = true;
            try
            {
                var cats = ShapeLibraryService.GetAllCategories();
                var names = new List<string> { "全部" };
                names.AddRange(cats.Where(c => !names.Contains(c)));
                if (!names.Contains(_categoryFilter)) _categoryFilter = "全部";

                stackCategory.Children.Clear();
                foreach (var cat in names)
                {
                    var rb = new RadioButton
                    {
                        Content = cat,
                        Tag = cat,
                        IsChecked = cat == _categoryFilter,
                        Style = (Style)FindResource("CategoryChip"),
                        GroupName = "cats",
                        Foreground = (Brush)FindResource("TextBrush")
                    };
                    rb.Checked += OnCategoryTabChanged;
                    stackCategory.Children.Add(rb);
                }
            }
            finally
            {
                _updatingUI = false;
            }

            // 过滤 + 排序
            var items = all.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(_filter))
            {
                string f = _filter;
                items = items.Where(m =>
                    m.Name.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.GetCategory().IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    string.Join(" ", m.GetTags()).IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            if (_categoryFilter != "全部")
            {
                items = items.Where(m => m.GetCategory() == _categoryFilter);
            }
            var list = items.ToList();
            switch (_sortMode)
            {
                case "最旧": list = list.OrderBy(m => m.CreatedAt).ToList(); break;
                case "名称": list = list.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase).ToList(); break;
                case "最近使用": list = list.OrderByDescending(m => m.LastUsedAt ?? "").ToList(); break;
                case "最常用": list = list.OrderByDescending(m => m.UseCount).ThenByDescending(m => m.LastUsedAt ?? "").ToList(); break;
                case "自定义": list = list.OrderBy(m => m.SortOrder).ThenByDescending(m => m.CreatedAt).ToList(); break;
                default: list = list.OrderByDescending(m => m.CreatedAt).ToList(); break;
            }
            if (list.Count > _loadedCount) list = list.GetRange(0, _loadedCount);

            // 构建卡片
            _cards.Clear();
            itemsControl.Items.Clear();
            if (list.Count == 0)
            {
                txtEmpty.Visibility = Visibility.Visible;
                txtEmpty.Text = (_filter != "" || _categoryFilter != "全部")
                    ? "没有符合筛选条件的收藏"
                    : "还没有收藏\n\n选中页面上的形状\n点击工具栏「收藏形状」或「收藏样式」";
            }
            else
            {
                txtEmpty.Visibility = Visibility.Collapsed;
                foreach (var meta in list)
                {
                    var card = CreateCard(meta);
                    _cards[meta.Id] = card;
                    itemsControl.Items.Add(card);
                }
            }

            lblStatus.Text = "共 " + list.Count + " 项" + (list.Count < all.Count ? " / " + all.Count : "");
            UpdateBatchUI();
        }

        // ---------------- 卡片 ----------------

        private Border CreateCard(LibraryMeta meta)
        {
            int cardW = 96, cardH = 114;

            var grid = new Grid { Width = cardW, Height = cardH };
            var rows = new RowDefinition { Height = new GridLength(86) };
            var row2 = new RowDefinition { Height = new GridLength(20) };
            grid.RowDefinitions.Add(rows);
            grid.RowDefinitions.Add(row2);

            // 缩略图
            var img = new Image
            {
                Stretch = Stretch.Uniform,
                Margin = new Thickness(5, 5, 5, 2),
                Source = LoadPreview(meta.Id)
            };
            Grid.SetRow(img, 0);

            // 名称
            var name = new TextBlock
            {
                Text = meta.Name,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontSize = 10.5,
                FontWeight = FontWeights.Medium,
                Foreground = (Brush)FindResource("TextBrush"),
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(6, 2, 6, 0)
            };
            Grid.SetRow(name, 1);

            // 徽章
            var badge = new Border
            {
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 5, 5, 0),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(5, 2, 5, 2),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(meta.Kind == "Style" ? "#7B1FA2" : "#0067C0")),
                Opacity = 0.9
            };
            var badgeText = new TextBlock
            {
                Text = meta.Kind == "Style" ? "样式" : "形状",
                FontSize = 9,
                Foreground = Brushes.White,
                FontFamily = new FontFamily("Microsoft YaHei UI")
            };
            badge.Child = badgeText;
            Grid.SetRow(badge, 0);

            grid.Children.Add(img);
            grid.Children.Add(name);
            grid.Children.Add(badge);

            var card = new Border
            {
                Style = (Style)FindResource("CardBase"),
                Child = grid,
                Tag = meta.Id,
                ToolTip = BuildTooltip(meta)
            };

            // hover 删除按钮（右上角，淡入淡出）
            var delBtn = new Button
            {
                Content = "\uE74D",
                Style = (Style)FindResource("IconButton"),
                Width = 22,
                Height = 22,
                FontSize = 10,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 5, 5, 0),
                Background = new SolidColorBrush(Color.FromRgb(0xE5, 0x48, 0x4D)),
                Foreground = Brushes.White,
                ToolTip = "删除",
                Opacity = 0,
                IsHitTestVisible = false,
                Cursor = Cursors.Hand
            };
            delBtn.Click += (s, e) => DeleteItem(meta.Id);
            Grid.SetRow(delBtn, 0);
            grid.Children.Add(delBtn);
            card.MouseEnter += (s, e) =>
            {
                delBtn.BeginAnimation(UIElement.OpacityProperty,
                    Anim(1.0, 0.15));
                delBtn.IsHitTestVisible = true;
            };
            card.MouseLeave += (s, e) =>
            {
                delBtn.BeginAnimation(UIElement.OpacityProperty,
                    Anim(0.0, 0.15));
                delBtn.IsHitTestVisible = false;
            };

            // 选中描边
            if (_selectedIds.Contains(meta.Id))
            {
                card.BorderBrush = (Brush)FindResource("AccentBrush");
                card.BorderThickness = new Thickness(2);
            }

            // 淡入动画（仅透明度，避免与样式 hover 上浮的 RenderTransform 冲突）
            card.Opacity = 0;
            card.Loaded += (s, e) =>
            {
                card.BeginAnimation(UIElement.OpacityProperty, Anim(1.0, 0.25));
            };

            // 事件
            card.MouseLeftButtonUp += (s, e) => OnCardClick(meta.Id);
            card.MouseLeftButtonDown += (s, e) => OnCardMouseDown(card, meta.Id, e);
            card.MouseMove += (s, e) => OnCardMouseMove(card, meta.Id, e);

            // 右键菜单
            var menu = new ContextMenu();
            menu.Items.Add(MenuItem("插入到当前页", () => ActivateItem(meta.Id, true)));
            menu.Items.Add(MenuItem("应用样式到选中形状", () => ShapeExporter.ApplyStyle(meta.Id)));
            menu.Items.Add(MenuItem("应用到本页同类形状", () => ShapeExporter.ApplyStyleToSameType(meta.Id)));
            menu.Items.Add(MenuItem("预览大图…", () => ShowPreview(meta.Id)));
            menu.Items.Add(new Separator());
            menu.Items.Add(MenuItem("移动到分类…", () => MoveToCategory(meta.Id)));
            menu.Items.Add(MenuItem("编辑标签…", () => EditTags(meta.Id)));
            menu.Items.Add(MenuItem("重命名", () => RenameItem(meta.Id)));
            menu.Items.Add(MenuItem("删除", () => DeleteItem(meta.Id)));
            menu.Items.Add(new Separator());
            menu.Items.Add(MenuItem("打开收藏文件夹", () => OpenFolder(meta.Id)));
            card.ContextMenu = menu;
            return card;
        }

        private static System.Windows.Media.Animation.DoubleAnimation Anim(double to, double seconds)
        {
            return new System.Windows.Media.Animation.DoubleAnimation(to, TimeSpan.FromSeconds(seconds))
            {
                EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
            };
        }

        private string BuildTooltip(LibraryMeta meta)
        {
            var tags = meta.GetTags();
            string t = meta.Name + "  [分类: " + meta.GetCategory() + "]";
            if (tags.Length > 0) t += "\n标签: " + string.Join("、", tags);
            if (meta.UseCount > 0) t += "\n使用 " + meta.UseCount + " 次";
            return t;
        }

        private MenuItem MenuItem(string header, Action action)
        {
            var mi = new MenuItem { Header = header, FontFamily = new FontFamily("Microsoft YaHei UI") };
            mi.Click += (s, e) => action();
            return mi;
        }

        private BitmapSource LoadPreview(string id)
        {
            BitmapSource cached;
            if (_previewCache.TryGetValue(id, out cached)) return cached;
            try
            {
                var bytes = ShapeLibraryService.LoadPreviewImage(id);
                if (bytes != null && bytes.Length > 0)
                {
                    using (var ms = new MemoryStream(bytes))
                    {
                        var bmp = new BitmapImage();
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = ms;
                        bmp.EndInit();
                        bmp.Freeze();
                        _previewCache[id] = bmp;
                        return bmp;
                    }
                }
            }
            catch { }
            return null;
        }

        private void OnCardClick(string id)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (_selectedIds.Contains(id)) _selectedIds.Remove(id);
                else _selectedIds.Add(id);
                RefreshSelectionVisual();
                UpdateBatchUI();
                return;
            }
            if (_selectedIds.Count > 0)
            {
                _selectedIds.Clear();
                RefreshSelectionVisual();
                UpdateBatchUI();
            }
            if (_settings.ClickToInsert)
            {
                ActivateItem(id);
            }
            else
            {
                ShowPreview(id);
            }
        }

        private Point _dragStart = new Point(-1, -1);
        private string _dragId = null;

        private void OnCardMouseDown(Border card, string id, MouseButtonEventArgs e)
        {
            _dragStart = e.GetPosition(card);
            _dragId = id;
        }

        private void OnCardMouseMove(Border card, string id, MouseEventArgs e)
        {
            if (_dragId == id && e.LeftButton == MouseButtonState.Pressed && _dragStart.X >= 0)
            {
                var pos = e.GetPosition(card);
                if (Math.Abs(pos.X - _dragStart.X) > 6 || Math.Abs(pos.Y - _dragStart.Y) > 6)
                {
                    _dragStart = new Point(-1, -1);
                    var data = new DataObject(typeof(string), id);
                    DragDrop.DoDragDrop(card, data, DragDropEffects.Move);
                }
            }
        }

        private void RefreshSelectionVisual()
        {
            foreach (var kv in _cards)
            {
                bool sel = _selectedIds.Contains(kv.Key);
                kv.Value.BorderBrush = sel ? (Brush)FindResource("AccentBrush") : (Brush)FindResource("BorderBrush");
                kv.Value.BorderThickness = sel ? new Thickness(2) : new Thickness(1);
            }
        }

        private void UpdateBatchUI()
        {
            panelBatch.Visibility = _selectedIds.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            lblSelected.Text = "已选 " + _selectedIds.Count + " 项";
        }

        private void ActivateItem(string id, bool forceInsert = false)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            if (meta.Kind == "Style" && !forceInsert)
            {
                ShapeExporter.ApplyStyle(id);
                ShapeLibraryService.TouchUsed(id);
            }
            else
            {
                if (ShapeExporter.InsertShape(id))
                {
                    ShapeLibraryService.TouchUsed(id);
                    ShowAdjustBar();
                }
            }
        }

        // ---------------- 调整条 / 批量 ----------------

        private void ShowAdjustBar()
        {
            panelAdjust.Visibility = Visibility.Visible;
        }

        private void OnAdjustClicked(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            string action = btn.Tag as string;
            if (action == "done")
            {
                panelAdjust.Visibility = Visibility.Collapsed;
                return;
            }
            ShapeExporter.AdjustSelected(action);
        }

        private void OnClearSelClicked(object sender, RoutedEventArgs e)
        {
            _selectedIds.Clear();
            RefreshSelectionVisual();
            UpdateBatchUI();
        }

        private void OnBatchDelClicked(object sender, RoutedEventArgs e)
        {
            if (_selectedIds.Count == 0) return;
            if (MessageBox.Show("确定删除选中的 " + _selectedIds.Count + " 项收藏？", "PPT形状库",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            foreach (var id in _selectedIds.ToList()) ShapeLibraryService.DeleteItem(id);
            _selectedIds.Clear();
            RefreshLibrary();
        }

        private void OnBatchMoveClicked(object sender, RoutedEventArgs e)
        {
            if (_selectedIds.Count == 0) return;
            var input = new InputDialogWpf("批量移动分类", "目标分类名称：", "");
            if (input.ShowDialog() != true || string.IsNullOrWhiteSpace(input.Value)) return;
            string cat = input.Value.Trim();
            foreach (var id in _selectedIds.ToList()) ShapeLibraryService.SetCategory(id, cat);
            _selectedIds.Clear();
            RefreshLibrary();
        }

        // ---------------- 事件 ----------------

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            txtPlaceholder.Visibility = string.IsNullOrEmpty(txtSearch.Text) ? Visibility.Visible : Visibility.Collapsed;
            _filter = txtSearch.Text.Trim();
            RefreshLibrary();
        }

        private void OnCategoryTabChanged(object sender, RoutedEventArgs e)
        {
            if (_updatingUI) return;
            var rb = sender as RadioButton;
            if (rb == null || rb.IsChecked != true) return;
            _categoryFilter = (rb.Tag ?? rb.Content).ToString();
            RefreshLibrary();
        }

        private void OnSortChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbSort.SelectedItem != null)
            {
                _sortMode = cmbSort.SelectedItem.ToString();
                RefreshLibrary();
            }
        }

        private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (e.VerticalOffset > e.ExtentHeight - 300)
            {
                int all = ShapeLibraryService.LoadAllItems().Count;
                if (_loadedCount < all)
                {
                    _loadedCount += _pageSize;
                    RefreshLibrary();
                }
            }
        }

        // ---------------- 管理 ----------------

        private void RenameItem(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            var input = new InputDialogWpf("重命名收藏", "名称：", meta.Name);
            if (input.ShowDialog() == true && !string.IsNullOrWhiteSpace(input.Value))
            {
                ShapeLibraryService.RenameItem(id, input.Value.Trim());
                RefreshLibrary();
            }
        }

        private void MoveToCategory(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            var input = new InputDialogWpf("移动到分类", "分类名称（输入新分类或已有分类）：", meta.GetCategory());
            if (input.ShowDialog() == true && !string.IsNullOrWhiteSpace(input.Value))
            {
                ShapeLibraryService.SetCategory(id, input.Value.Trim());
                RefreshLibrary();
            }
        }

        private void EditTags(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            var input = new InputDialogWpf("编辑标签", "标签（多个用逗号分隔）：", meta.Tags ?? "");
            if (input.ShowDialog() == true)
            {
                ShapeLibraryService.SetTags(id, input.Value);
                RefreshLibrary();
            }
        }

        private void DeleteItem(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            if (MessageBox.Show("确定删除收藏「" + meta.Name + "」？", "PPT形状库",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            ShapeLibraryService.DeleteItem(id);
            BitmapSource tmp;
            if (_previewCache.TryGetValue(id, out tmp))
            {
                _previewCache.Remove(id);
            }
            RefreshLibrary();
        }

        private void ShowPreview(string id)
        {
            var meta = ShapeLibraryService.LoadMeta(id);
            if (meta == null) return;
            var img = LoadPreview(id);
            var win = new Window
            {
                Title = meta.Name,
                Width = 460,
                Height = 400,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = (Brush)FindResource("BgBrush")
            };
            var content = new Grid { Margin = new Thickness(16) };
            var preview = new Image
            {
                Source = img,
                Stretch = Stretch.Uniform
            };
            var tip = new TextBlock
            {
                Text = meta.Kind == "Style" ? "样式收藏" : "形状收藏",
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 8, 0, 0),
                Foreground = (Brush)FindResource("SubTextBrush"),
                FontFamily = new FontFamily("Microsoft YaHei UI")
            };
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(preview, 0);
            Grid.SetRow(tip, 1);
            grid.Children.Add(preview);
            grid.Children.Add(tip);
            content.Children.Add(grid);
            win.Content = content;
            win.ShowDialog();
        }

        private static void OpenFolder(string id)
        {
            string dir = Path.Combine(ShapeLibraryService.ItemsDir, id);
            System.Diagnostics.Process.Start("explorer.exe", dir);
        }
    }

    /// <summary>WPF 版单行输入对话框。</summary>
    public class InputDialogWpf : Window
    {
        private readonly System.Windows.Controls.TextBox _textBox;

        public string Value
        {
            get { return _textBox.Text; }
        }

        public InputDialogWpf(string title, string label, string initialValue = "")
        {
            Title = title;
            Width = 360;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;

            var grid = new Grid { Margin = new Thickness(16) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var lbl = new TextBlock
            {
                Text = label,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                Margin = new Thickness(0, 0, 0, 8)
            };
            _textBox = new System.Windows.Controls.TextBox
            {
                Text = initialValue,
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                Padding = new Thickness(4),
                Height = 26
            };
            var btns = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            var ok = new Button { Content = "确定", Width = 80, Height = 28, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancel = new Button { Content = "取消", Width = 80, Height = 28, IsCancel = true };
            ok.Click += (s, e) => { DialogResult = true; Close(); };
            cancel.Click += (s, e) => { DialogResult = false; Close(); };
            btns.Children.Add(ok);
            btns.Children.Add(cancel);

            Grid.SetRow(lbl, 0);
            Grid.SetRow(_textBox, 1);
            Grid.SetRow(btns, 2);
            grid.Children.Add(lbl);
            grid.Children.Add(_textBox);
            grid.Children.Add(btns);
            Content = grid;
            Loaded += (s, e) => _textBox.Focus();
        }
    }

    /// <summary>WPF 面板的 WinForms 宿主（VSTO 任务窗格仅接受 UserControl）。降级路径，代码冻结。</summary>
    public class WpfPaneHost : System.Windows.Forms.UserControl, ILibraryPane
    {
        public readonly ShapeLibraryPaneWpf WpfPane;

        public WpfPaneHost(ShapeLibraryPaneWpf pane)
        {
            WpfPane = pane;
            var host = new System.Windows.Forms.Integration.ElementHost
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Child = pane
            };
            Controls.Add(host);
        }

        public void RefreshLibrary()
        {
            WpfPane.RefreshLibrary();
        }

        public void ApplySettings(Settings s)
        {
            WpfPane.ApplySettings(s);
        }
    }
}
