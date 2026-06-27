using System;
using System.Collections.Generic;
using System.Linq;

namespace UpslideClone.Core.Design
{
    /// <summary>A shape's bounding box (PowerPoint points). Pure value type.</summary>
    public struct LayoutBox
    {
        public float Left, Top, Width, Height;
        public LayoutBox(float left, float top, float width, float height) { Left = left; Top = top; Width = width; Height = height; }
        public float Right => Left + Width;
        public float Bottom => Top + Height;
        public float CenterX => Left + Width / 2f;
        public float CenterY => Top + Height / 2f;
    }

    public enum AlignMode { Left, CenterHorizontal, Right, Top, Middle, Bottom }
    public enum DistributeAxis { Horizontal, Vertical }
    public enum SizeMatch { Width, Height, Both }

    /// <summary>
    /// Pure shape-layout geometry for the PowerPoint design tools (Smart Align,
    /// Resize &amp; Distribute). No Office dependency → unit-testable. Each method
    /// returns new boxes in the SAME order as the input; sizes are preserved
    /// except in <see cref="MatchSize"/>.
    /// </summary>
    public static class AlignEngine
    {
        /// <summary>Align all boxes to the selection's bounding box (Upslide "Smart Align").</summary>
        public static List<LayoutBox> Align(IList<LayoutBox> boxes, AlignMode mode)
        {
            if (boxes == null) throw new ArgumentNullException(nameof(boxes));
            var result = boxes.ToList();
            if (result.Count == 0) return result;

            float minL = result.Min(b => b.Left);
            float minT = result.Min(b => b.Top);
            float maxR = result.Max(b => b.Right);
            float maxB = result.Max(b => b.Bottom);
            float cx = (minL + maxR) / 2f;
            float cy = (minT + maxB) / 2f;

            for (int i = 0; i < result.Count; i++)
            {
                var b = result[i];
                switch (mode)
                {
                    case AlignMode.Left: b.Left = minL; break;
                    case AlignMode.Right: b.Left = maxR - b.Width; break;
                    case AlignMode.CenterHorizontal: b.Left = cx - b.Width / 2f; break;
                    case AlignMode.Top: b.Top = minT; break;
                    case AlignMode.Bottom: b.Top = maxB - b.Height; break;
                    case AlignMode.Middle: b.Top = cy - b.Height / 2f; break;
                }
                result[i] = b;
            }
            return result;
        }

        /// <summary>Distribute boxes so the gaps between them are equal; ends stay fixed.</summary>
        public static List<LayoutBox> Distribute(IList<LayoutBox> boxes, DistributeAxis axis)
        {
            if (boxes == null) throw new ArgumentNullException(nameof(boxes));
            var result = boxes.ToList();
            if (result.Count < 3) return result; // nothing to distribute between 2 ends

            var order = Enumerable.Range(0, result.Count)
                .OrderBy(i => axis == DistributeAxis.Horizontal ? result[i].Left : result[i].Top)
                .ToList();

            float spanStart = axis == DistributeAxis.Horizontal ? result[order[0]].Left : result[order[0]].Top;
            float spanEnd = axis == DistributeAxis.Horizontal ? result[order[order.Count - 1]].Right : result[order[order.Count - 1]].Bottom;
            float sumSize = order.Sum(i => axis == DistributeAxis.Horizontal ? result[i].Width : result[i].Height);
            float gap = (spanEnd - spanStart - sumSize) / (order.Count - 1);

            float cursor = spanStart;
            foreach (int i in order)
            {
                var b = result[i];
                if (axis == DistributeAxis.Horizontal) { b.Left = cursor; cursor += b.Width + gap; }
                else { b.Top = cursor; cursor += b.Height + gap; }
                result[i] = b;
            }
            return result;
        }

        /// <summary>Resize every box to match the first box's size (Upslide "same size").</summary>
        public static List<LayoutBox> MatchSize(IList<LayoutBox> boxes, SizeMatch which)
        {
            if (boxes == null) throw new ArgumentNullException(nameof(boxes));
            var result = boxes.ToList();
            if (result.Count < 2) return result;

            var refBox = result[0];
            for (int i = 1; i < result.Count; i++)
            {
                var b = result[i];
                if (which == SizeMatch.Width || which == SizeMatch.Both) b.Width = refBox.Width;
                if (which == SizeMatch.Height || which == SizeMatch.Both) b.Height = refBox.Height;
                result[i] = b;
            }
            return result;
        }
    }
}
