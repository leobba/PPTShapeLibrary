using System;
using System.Globalization;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTShapeLibrary
{
    /// <summary>
    /// 从选中形状提取样式属性（StylePreset），或把样式应用到选中形状。
    /// 所有取值均带 try/catch，某个属性不支持时跳过，不影响整体。
    /// </summary>
    public static class ShapeStyleExtractor
    {
        // ---------------- 提取 ----------------

        public static StylePreset Extract(PowerPoint.ShapeRange selection, string name)
        {
            var preset = new StylePreset
            {
                Name = name,
                CreatedAt = DateTime.Now
            };

            if (selection == null || selection.Count == 0) return preset;

            PowerPoint.Shape shape = selection[1];

            Try(() =>
            {
                preset.FillType = (int)shape.Fill.Type;
                preset.FillColor = RgbToHex(shape.Fill.ForeColor.RGB);
                preset.FillTransparency = shape.Fill.Transparency;
            });
            Try(() =>
            {
                if (shape.Fill.GradientStops.Count > 0)
                {
                    preset.GradientStops = new System.Collections.Generic.List<GradientStop>();
                    for (int i = 1; i <= shape.Fill.GradientStops.Count; i++)
                    {
                        var stop = shape.Fill.GradientStops[i];
                        preset.GradientStops.Add(new GradientStop
                        {
                            Position = stop.Position,
                            Color = RgbToHex(stop.Color.RGB),
                            Transparency = stop.Transparency
                        });
                    }
                }
            });
            Try(() => preset.GradientAngle = (int)shape.Fill.GradientAngle);
            Try(() => preset.GradientStyle = (int)shape.Fill.GradientStyle);
            Try(() => preset.GradientVariant = shape.Fill.GradientVariant);

            Try(() =>
            {
                preset.LineVisible = shape.Line.Visible == MsoTriState.msoTrue;
                preset.LineColor = RgbToHex(shape.Line.ForeColor.RGB);
                preset.LineWeight = shape.Line.Weight;
                preset.LineDashStyle = (int)shape.Line.DashStyle;
                preset.LineTransparency = shape.Line.Transparency;
            });

            Try(() =>
            {
                var shadow = shape.Shadow;
                preset.ShadowVisible = shadow.Visible == MsoTriState.msoTrue;
                preset.ShadowColor = RgbToHex(shadow.ForeColor.RGB);
                preset.ShadowTransparency = shadow.Transparency;
                preset.ShadowOffsetX = shadow.OffsetX;
                preset.ShadowOffsetY = shadow.OffsetY;
                preset.ShadowSize = shadow.Size;
            });

            Try(() =>
            {
                var effect = shape.Shadow;
                preset.ShadowBlur = effect.Blur;
                preset.ShadowTransparency = effect.Transparency;
            });

            Try(() =>
            {
                var td = shape.ThreeD;
                preset.ThreeDVisible = td.Visible == MsoTriState.msoTrue;
                preset.ThreeDColor = RgbToHex(td.ExtrusionColor.RGB);
                preset.ThreeDDepth = td.Depth;
                preset.ThreeDContourWidth = td.ContourWidth;
                preset.ThreeDContourColor = RgbToHex(td.ContourColor.RGB);
            });

            Try(() =>
            {
                var tf = shape.TextFrame2;
                if (tf != null && tf.HasText == MsoTriState.msoTrue && tf.TextRange.Length > 0)
                {
                    var font = tf.TextRange.Font;
                    preset.FontName = font.Name;
                    preset.FontSize = font.Size;
                    preset.FontBold = font.Bold == MsoTriState.msoTrue;
                    preset.FontItalic = font.Italic == MsoTriState.msoTrue;
                    preset.FontColor = RgbToHex(font.Fill.ForeColor.RGB);
                    preset.ParagraphAlignment = (int)tf.TextRange.ParagraphFormat.Alignment;
                    preset.TextVerticalAnchor = (int)tf.VerticalAnchor;
                    preset.TextLineSpacing = tf.TextRange.ParagraphFormat.SpaceWithin;
                    preset.TextMarginLeft = tf.MarginLeft;
                    preset.TextMarginRight = tf.MarginRight;
                    preset.TextMarginTop = tf.MarginTop;
                    preset.TextMarginBottom = tf.MarginBottom;
                }
            });

            Try(() =>
            {
                preset.Rotation = shape.Rotation;
            });
            Try(() =>
            {
                if (shape.Adjustments.Count >= 1)
                {
                    preset.CornerRadius = shape.Adjustments[1];
                }
            });

            Try(() =>
            {
                var refl = shape.Reflection;
                preset.ReflectionVisible = refl.Type != MsoReflectionType.msoReflectionTypeNone;
            });
            Try(() =>
            {
                var glow = shape.Glow;
                preset.GlowVisible = glow.Radius > 0;
                preset.GlowColor = RgbToHex(glow.Color.RGB);
                preset.GlowRadius = glow.Radius;
            });
            Try(() =>
            {
                var soft = shape.SoftEdge;
                preset.SoftEdgeVisible = soft.Type != MsoSoftEdgeType.msoSoftEdgeTypeNone;
                preset.SoftEdgeRadius = soft.Radius;
            });

            return preset;
        }

        // ---------------- 应用 ----------------

        public static void Apply(PowerPoint.ShapeRange selection, StylePreset preset)
        {
            if (selection == null || selection.Count == 0 || preset == null) return;

            for (int i = 1; i <= selection.Count; i++)
            {
                ApplyToShape(selection[i], preset);
            }
        }

        private static void ApplyToShape(PowerPoint.Shape shape, StylePreset p)
        {
            Try(() =>
            {
                if (p.FillType != null)
                {
                    if (p.GradientStops != null && p.GradientStops.Count >= 2)
                    {
                        MsoGradientStyle style = (MsoGradientStyle)(p.GradientStyle ?? (int)MsoGradientStyle.msoGradientHorizontal);
                        int variant = p.GradientVariant ?? 1;
                        shape.Fill.TwoColorGradient(style, variant);
                        var stops = shape.Fill.GradientStops;
                        while (stops.Count > p.GradientStops.Count) stops.Delete(2);
                        for (int si = 0; si < p.GradientStops.Count; si++)
                        {
                            if (si < stops.Count) stops[si + 1].Color.RGB = HexToRgb(p.GradientStops[si].Color);
                            else stops.Insert(RgbToInt(HexToRgb(p.GradientStops[si].Color)), (float)p.GradientStops[si].Position);
                            stops[si + 1].Position = (float)p.GradientStops[si].Position;
                            if (p.GradientStops[si].Transparency != null)
                            {
                                stops[si + 1].Transparency = (float)p.GradientStops[si].Transparency;
                            }
                        }
                        Try(() => shape.Fill.GradientAngle = (float)(p.GradientAngle ?? 90));
                    }
                    else if (p.GradientStops != null && p.GradientStops.Count == 1)
                    {
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = HexToRgb(p.GradientStops[0].Color);
                    }
                    else if (p.FillColor != null)
                    {
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = HexToRgb(p.FillColor);
                    }
                    if (p.FillTransparency != null) shape.Fill.Transparency = (float)p.FillTransparency;
                }
            });

            Try(() =>
            {
                if (p.LineVisible == true)
                {
                    shape.Line.Visible = MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = HexToRgb(p.LineColor ?? "000000");
                    if (p.LineWeight != null) shape.Line.Weight = (float)p.LineWeight;
                    if (p.LineDashStyle != null) shape.Line.DashStyle = (MsoLineDashStyle)p.LineDashStyle;
                    if (p.LineTransparency != null) shape.Line.Transparency = (float)p.LineTransparency;
                }
                else if (p.LineVisible == false)
                {
                    shape.Line.Visible = MsoTriState.msoFalse;
                }
            });

            Try(() =>
            {
                if (p.ShadowVisible != null)
                {
                    shape.Shadow.Visible = p.ShadowVisible.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                }
                if (p.ShadowColor != null) shape.Shadow.ForeColor.RGB = HexToRgb(p.ShadowColor);
                if (p.ShadowTransparency != null) shape.Shadow.Transparency = (float)p.ShadowTransparency;
                if (p.ShadowOffsetX != null) shape.Shadow.OffsetX = (float)p.ShadowOffsetX;
                if (p.ShadowOffsetY != null) shape.Shadow.OffsetY = (float)p.ShadowOffsetY;
                if (p.ShadowBlur != null) shape.Shadow.Blur = (float)p.ShadowBlur;
                if (p.ShadowSize != null) shape.Shadow.Size = (float)p.ShadowSize;
            });

            Try(() =>
            {
                if (p.ThreeDVisible != null)
                {
                    shape.ThreeD.Visible = p.ThreeDVisible.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                }
                if (p.ThreeDColor != null) shape.ThreeD.ExtrusionColor.RGB = HexToRgb(p.ThreeDColor);
                if (p.ThreeDDepth != null) shape.ThreeD.Depth = (float)p.ThreeDDepth;
                if (p.ThreeDContourWidth != null) shape.ThreeD.ContourWidth = (float)p.ThreeDContourWidth;
                if (p.ThreeDContourColor != null) shape.ThreeD.ContourColor.RGB = HexToRgb(p.ThreeDContourColor);
            });

            Try(() =>
            {
                var tf = shape.TextFrame2;
                if (tf != null && tf.TextRange.Length > 0)
                {
                    if (p.FontName != null) tf.TextRange.Font.Name = p.FontName;
                    if (p.FontSize != null) tf.TextRange.Font.Size = (float)p.FontSize;
                    if (p.FontBold != null) tf.TextRange.Font.Bold = p.FontBold.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                    if (p.FontItalic != null) tf.TextRange.Font.Italic = p.FontItalic.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                    if (p.FontColor != null) tf.TextRange.Font.Fill.ForeColor.RGB = HexToRgb(p.FontColor);
                    if (p.ParagraphAlignment != null) tf.TextRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)p.ParagraphAlignment;
                    if (p.TextLineSpacing != null) tf.TextRange.ParagraphFormat.SpaceWithin = (float)p.TextLineSpacing;
                }
                if (p.TextVerticalAnchor != null) tf.VerticalAnchor = (MsoVerticalAnchor)p.TextVerticalAnchor;
                if (p.TextMarginLeft != null) tf.MarginLeft = (float)p.TextMarginLeft;
                if (p.TextMarginRight != null) tf.MarginRight = (float)p.TextMarginRight;
                if (p.TextMarginTop != null) tf.MarginTop = (float)p.TextMarginTop;
                if (p.TextMarginBottom != null) tf.MarginBottom = (float)p.TextMarginBottom;
            });

            Try(() =>
            {
                if (p.Rotation != null) shape.Rotation = (float)p.Rotation;
            });
            Try(() =>
            {
                if (p.CornerRadius != null && shape.Adjustments.Count >= 1)
                {
                    shape.Adjustments[1] = (float)p.CornerRadius;
                }
            });

            Try(() =>
            {
                if (p.ReflectionVisible != null)
                {
                    shape.Reflection.Type = p.ReflectionVisible.Value
                        ? MsoReflectionType.msoReflectionType5
                        : MsoReflectionType.msoReflectionTypeNone;
                }
            });
            Try(() =>
            {
                if (p.GlowVisible != null) shape.Glow.Radius = p.GlowVisible.Value ? (float)(p.GlowRadius ?? 8) : 0f;
                if (p.GlowColor != null) shape.Glow.Color.RGB = HexToRgb(p.GlowColor);
                if (p.GlowRadius != null) shape.Glow.Radius = (float)p.GlowRadius;
            });
            Try(() =>
            {
                if (p.SoftEdgeVisible != null)
                {
                    shape.SoftEdge.Type = p.SoftEdgeVisible.Value
                        ? MsoSoftEdgeType.msoSoftEdgeType1
                        : MsoSoftEdgeType.msoSoftEdgeTypeNone;
                }
                if (p.SoftEdgeRadius != null) shape.SoftEdge.Radius = (float)p.SoftEdgeRadius;
            });
        }

        // ---------------- 序列化 ----------------

        public static string ToJson(StylePreset preset)
        {
            var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
            return ser.Serialize(preset);
        }

        public static StylePreset FromJson(string json)
        {
            if (string.IsNullOrEmpty(json)) return null;
            try
            {
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                return ser.Deserialize<StylePreset>(json);
            }
            catch
            {
                return null;
            }
        }

        // ---------------- 工具 ----------------

        public static string RgbToHex(int rgb)
        {
            int r = (rgb >> 0) & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        public static int HexToRgb(string hex)
        {
            if (string.IsNullOrEmpty(hex)) hex = "000000";
            hex = hex.TrimStart('#');
            int r = int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber);
            int g = int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber);
            int b = int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber);
            return r | (g << 8) | (b << 16);
        }

        private static int RgbToInt(int rgb)
        {
            int r = (rgb >> 0) & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return (r << 16) | (g << 8) | b;
        }

        private static void Try(Action action)
        {
            try { action(); }
            catch { /* 单属性失败不影响整体 */ }
        }
    }
}

