using System;
using System.Collections.Generic;

namespace PPTShapeLibrary
{
    /// <summary>形状样式快照（JSON 可序列化，用于收藏/应用样式）</summary>
    [Serializable]
    public class StylePreset
    {
        public string Name { get; set; }
        public DateTime CreatedAt { get; set; }

        // ---- 填充 ----
        public int? FillType { get; set; }                 // MsoFillType
        public string FillColor { get; set; }             // "RRGGBB"（无 #）
        public double? FillTransparency { get; set; }     // 0-1
        public List<GradientStop> GradientStops { get; set; }
        public int? GradientAngle { get; set; }           // 渐变方向角度（线性渐变）
        public int? GradientStyle { get; set; }           // MsoGradientStyle（射线/线性等）
        public int? GradientVariant { get; set; }         // 1-4

        // ---- 线条 ----
        public bool? LineVisible { get; set; }
        public string LineColor { get; set; }
        public double? LineWeight { get; set; }
        public int? LineDashStyle { get; set; }           // MsoLineDashStyle
        public double? LineTransparency { get; set; }     // 0-1

        // ---- 阴影 ----
        public bool? ShadowVisible { get; set; }
        public string ShadowColor { get; set; }
        public double? ShadowTransparency { get; set; }
        public double? ShadowBlur { get; set; }
        public double? ShadowOffsetX { get; set; }
        public double? ShadowOffsetY { get; set; }
        public double? ShadowSize { get; set; }

        // ---- 三维 ----
        public bool? ThreeDVisible { get; set; }
        public string ThreeDColor { get; set; }
        public double? ThreeDDepth { get; set; }
        public double? ThreeDContourWidth { get; set; }
        public string ThreeDContourColor { get; set; }

        // ---- 文本 ----
        public string FontName { get; set; }
        public double? FontSize { get; set; }
        public bool? FontBold { get; set; }
        public bool? FontItalic { get; set; }
        public string FontColor { get; set; }
        public int? ParagraphAlignment { get; set; }      // MsoParagraphAlignment
        public int? TextVerticalAnchor { get; set; }      // MsoVerticalAnchor
        public double? TextMarginLeft { get; set; }
        public double? TextMarginRight { get; set; }
        public double? TextMarginTop { get; set; }
        public double? TextMarginBottom { get; set; }
        public double? TextLineSpacing { get; set; }      // SpaceWithin

        // ---- 形状变换 ----
        public double? Rotation { get; set; }             // 旋转角度
        public double? CornerRadius { get; set; }         // 圆角矩形 Adjustments[1]

        // ---- 效果（反射/发光/柔化边缘）----
        public bool? ReflectionVisible { get; set; }
        public bool? GlowVisible { get; set; }
        public string GlowColor { get; set; }
        public double? GlowRadius { get; set; }
        public bool? SoftEdgeVisible { get; set; }
        public double? SoftEdgeRadius { get; set; }
    }

    [Serializable]
    public class GradientStop
    {
        public double Position { get; set; }
        public string Color { get; set; }
        public double? Transparency { get; set; }
    }
}
