using System;

namespace TemplateCooking.Domain.Layout
{
    public class LayoutElement
    {
        public LayoutElement(RcPosition topLeft, RcDimensions area)
        {
            TopLeft = topLeft ?? throw new ArgumentNullException(nameof(topLeft));
            Area = area ?? throw new ArgumentNullException(nameof(area));
        }

        public RcPosition TopLeft { get; }
        public RcDimensions Area { get; }
    }
}