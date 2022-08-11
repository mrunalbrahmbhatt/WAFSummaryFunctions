using CsvHelper.Configuration;

namespace WAFSummaryApps.Models
{
    public class WAFRecord
    {
        public string Category { get; set; }
        public string LinkText { get; set; }
        public string Link { get; set; }
        public string Priority { get; set; }
        public string ReportingCategory { get; set; }
        public string ReportingSubcategory { get; set; }
        public int Weight { get; set; }
        public string Context { get; set; }
    }

    public class WAFRecordClassMap : ClassMap<WAFRecord>
    {
        public WAFRecordClassMap()
        {
            Map(m => m.Category).Name("Category");
            Map(m => m.LinkText).Name("Link-Text");
            Map(m => m.Link).Name("Link");
            Map(m => m.Priority).Name("Priority");
            Map(m => m.ReportingCategory).Name("ReportingCategory");
            Map(m => m.ReportingSubcategory).Name("ReportingSubcategory");
            Map(m => m.Weight).Name("Weight");
            Map(m => m.Context).Name("Context");
        }
    }

}
