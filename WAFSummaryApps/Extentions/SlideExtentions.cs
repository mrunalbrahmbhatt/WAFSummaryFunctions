using ShapeCrawler;

namespace WAFSummaryApps.Extentions
{
    public static class SlideExtentions
    {
        public static IAutoShape AsAutoShape(this ISlide slide, int slideNumber)
        {
            return (IAutoShape)slide.Shapes[slideNumber];
        }
    }
}
