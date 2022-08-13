using CsvHelper;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using ShapeCrawler;
using ShapeCrawler.Shapes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WAFSummaryApps.Extentions;
using WAFSummaryApps.Models;

namespace WAFSummaryApps
{
    public class WAFSummaryFunctions
    {
        [FunctionName("PopulatePPT")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] BlobListItem csvBlobListItem,
            [Blob("csv/{name}", FileAccess.Read)] Stream csvWAF,
            [Blob("template/PnP_PowerPointReport_Template.pptx", FileAccess.Read)] Stream pptTemplate,
            [Blob("template/WAF Category Descriptions.csv", FileAccess.Read)] Stream csvDescription,
            IBinder pptBinder,
            ILogger log)
        {
            Dictionary<string, PillarInfo> pillarInfoList = null;
            var outputBlobName = $"WAFSummary{DateTime.Today:ddMMyyyyHHss}.pptx";

            var categoryDescriptions = GetCategoryDescription(csvDescription);

            string fileContent = GetContent(csvWAF);
            var wafRecords = GetWafRecords(fileContent);
            wafRecords.ForEach(r =>
            {
                if (string.IsNullOrWhiteSpace(r.ReportingCategory))
                {
                    r.ReportingCategory = "Uncategorized";
                }
            });
            var pillars = categoryDescriptions.Select(r => r.Pillar).Distinct().ToList();

            //Init Pillar Info
            pillarInfoList = InitPillarInfo(pillars);


            string overallScore = string.Empty;
            GetPillarScore(pillarInfoList, fileContent, out overallScore);
            GetPillarDescription(pillarInfoList, categoryDescriptions, pillars);

            //Dulication
            IPresentation presentation = null;
            IPresentation template = null;

            MemoryStream memPresentation = new MemoryStream();
            MemoryStream memTemplate = new MemoryStream();
            MemoryStream memFinalPresentation = new MemoryStream();

            try
            {
                //Create the streams.
                pptTemplate.Position = 0;
                pptTemplate.CopyTo(memPresentation);
                pptTemplate.Position = 0;
                pptTemplate.CopyTo(memTemplate);

                memPresentation.Position = 0;
                memTemplate.Position = 0;

                template = SCPresentation.Open(memTemplate, true);
                presentation = SCPresentation.Open(memPresentation, true);


                var titleSlide = template.Slides[7];
                var summarySlide = template.Slides[8];
                var detailSlide = template.Slides[9];
                var endSlide = template.Slides[10];

                var title = "Well-Architected [pillar] Assessment";
                var reportDate = DateTime.Today.ToString("yyyy-MM-dd-HHmm");
                var localReportDate = DateTime.Today.ToString("g");

                foreach (var pillar in pillars)
                {
                    var pillarData = wafRecords.Where(r => r.Category.Equals(pillar)).ToList();
                    var pillarInfo = pillarInfoList[pillar];

                    var slideTitle = title.Replace("[pillar]", pillar.Substring(0, 1).ToUpper() + pillar.Substring(1).ToLower());

                    presentation.Slides.Add(titleSlide);
                    ISlide newTitleSlide = presentation.Slides[presentation.Slides.Count - 1];
                    newTitleSlide.AsAutoShape(2).TextBox.Text = slideTitle;
                    newTitleSlide.AsAutoShape(3).TextBox.Text = newTitleSlide.AsAutoShape(3).TextBox.Text.Replace("[Report_Date]", localReportDate);

                    //Add logic to get overall score
                    presentation.Slides.Add(summarySlide);
                    ISlide newSummarySlide = presentation.Slides[presentation.Slides.Count - 1];

                    newSummarySlide.AsAutoShape(2).TextBox.Text = pillarInfo.Score.ToZeroIfNullorEmpty();
                    newSummarySlide.AsAutoShape(3).TextBox.Text = pillarInfo.Description;
                    var summBarScore = int.Parse(pillarInfo.Score.ToZeroIfNullorEmpty()) * 3.35 + 72;
                    newSummarySlide.Shapes[10].X = (int)summBarScore;

                    List<CategoryListItem> categoryListItems = new List<CategoryListItem>();

                    var categorys = pillarData
                        .OrderByDescending(p => p.Weight)
                        .Select(p => p.ReportingCategory)
                        .Distinct()
                        .ToList();

                    foreach (var category in categorys)
                    {
                        var weight = pillarData
                            .Where(p => p.ReportingCategory.Equals(category))
                            .Select(p => p.Weight).Sum();
                        var count = pillarData
                            .Where(p => p.ReportingCategory.Equals(category))
                            .Count();
                        var score = weight / count;

                        var weightiestCount = pillarData
                            .Where(p => p.ReportingCategory.Equals(category) && weight >= 0) //MinimumReportLevel is missing
                            .Count();

                        categoryListItems.Add(new CategoryListItem
                        {
                            Category = category,
                            Score = score,
                            WeightCount = weightiestCount
                        });

                    }

                    categoryListItems = categoryListItems.OrderByDescending(c => c.Score).ToList();

                    int counter = 13;//Shape count for the slide to start adding scores
                    var categoryCounter = 0;
                    var areaIconX = 378.1129;
                    var areaIconY = new double[] { 176.4359, 217.6319, 258.3682, 299.1754, 339.8692, 382.6667, 423.9795, 461.0491 };



                    foreach (var category in categoryListItems)
                    {
                        if (category.Category != "Uncategorized")
                        {
                            try
                            {
                                //$newSummarySlide.Shapes[8] #Domain 1 Icon
                                newSummarySlide.AsAutoShape(counter).TextBox.Text = category.WeightCount.ToString("#");
                                newSummarySlide.AsAutoShape(counter + 1).TextBox.Text = category.Category;
                                counter += 3;
                                IShape categoryShape = null;


                                if (category.Score < 33)
                                {
                                    categoryShape = summarySlide.Shapes[37];
                                }
                                else if (category.Score > 33 && category.Score < 67)
                                {
                                    categoryShape = summarySlide.Shapes[38];
                                }
                                else if (category.Score > 67)
                                {
                                    categoryShape = summarySlide.Shapes[39];
                                }
                                else
                                {
                                    categoryShape = summarySlide.Shapes[38];
                                }

                                newSummarySlide.Shapes.Append(categoryShape);
                                var newShapeIndex = newSummarySlide.Shapes.Count - 1;
                                newSummarySlide.Shapes[newShapeIndex].Y = (int)areaIconX;
                                newSummarySlide.Shapes[newShapeIndex].X = (int)areaIconY[categoryCounter];
                                categoryCounter++;
                            }
                            catch { }
                        }
                    }

                    break;
                }
                var outboundBlob = new BlobAttribute($"powerpoint/{outputBlobName}", FileAccess.Write);

                presentation.SaveAs(memFinalPresentation);


                using (var writer = pptBinder.Bind<Stream>(outboundBlob))
                {
                    memFinalPresentation.Position = 0;
                    writer.Write(memFinalPresentation.ToArray());
                };

            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {

                presentation.Close();
                template.Close();
                memPresentation.Dispose();
                memTemplate.Dispose();
                memFinalPresentation.Dispose();
            }

            log.LogInformation("C# HTTP trigger function processed a request.");
            return new OkObjectResult($"powerpoint/{outputBlobName}");
        }



        private void GetPillarDescription(Dictionary<string, PillarInfo> pillarInfo, IEnumerable<WafCategoryDescription> categoryDescriptions, List<string> pillars)
        {
            foreach (var pillar in pillars)
            {
                pillarInfo[pillar].Description = categoryDescriptions
                  .Where(d => d.Pillar.Contains(pillar) && d.Category.Equals("Survey Level Group")).Select(d => d.Description).FirstOrDefault();
            }
        }

        private Dictionary<string, PillarInfo> InitPillarInfo(List<string> pillars)
        {
            Dictionary<string, PillarInfo> pillarInfo = new Dictionary<string, PillarInfo>();
            foreach (var pillar in pillars)
            {
                pillarInfo.Add(pillar, new PillarInfo() { Name = pillar });
            }
            return pillarInfo;
        }

        private void GetPillarScore(Dictionary<string, PillarInfo> pillarInfo, string fileContent, out string overallScore)
        {
            var fileContentArray = fileContent.Split("\n");
            overallScore = string.Empty;
            for (int i = 3; i < 8; i++)
            {
                if (fileContentArray[i].Contains("overall"))
                {
                    overallScore = fileContentArray[i].Split(',')[2].Trim(new char[] { '\'' }).Split('/')[0];
                }
                if (fileContentArray[i].Contains("Cost Optimization"))
                {
                    pillarInfo["Cost Optimization"].Score = fileContentArray[i].Split(',')[2].Trim('\'').Split('/')[0];
                }
                if (fileContentArray[i].Contains("Reliability"))
                {
                    pillarInfo["Reliability"].Score = fileContentArray[i].Split(',')[2].Trim('\'').Split('/')[0];
                }
                if (fileContentArray[i].Contains("Operational Excellence"))
                {
                    pillarInfo["Operational Excellence"].Score = fileContentArray[i].Split(',')[2].Trim('\'').Split('/')[0];

                }
                if (fileContentArray[i].Contains("Performance Efficiency"))
                {
                    pillarInfo["Performance Efficiency"].Score = fileContentArray[i].Split(',')[2].Trim('\'').Split('/')[0];
                }
                if (fileContentArray[i].Contains("Security"))
                {
                    pillarInfo["Security"].Score = fileContentArray[i].Split(',')[2].Trim('\'').Split('/')[0];
                }
                if (fileContentArray[i].Equals(",,,,,"))
                {
                    // End early if not all pillars assessed
                    break;
                }
            }
        }

        private string GetContent(Stream fileStream)
        {
            using var reader = new StreamReader(fileStream, System.Text.Encoding.ASCII);
            return reader.ReadToEnd();
        }



        private List<WAFRecord> GetWafRecords(string csvContent)
        {
            int start = csvContent.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context");
            int end = csvContent.IndexOf("-----------,,,,,");
            string test = csvContent[start..end];
            using TextReader reader = new StringReader(test);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            {
                csv.Context.RegisterClassMap<WAFRecordClassMap>();
                return csv.GetRecords<WAFRecord>().ToList();
            }
        }

        private List<WafCategoryDescription> GetCategoryDescription(Stream csvDescription)
        {
            //IEnumerable<WafCategoryDescription> categoryDescriptions;
            using (var reader = new StreamReader(csvDescription))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                return csv.GetRecords<WafCategoryDescription>().ToList();
            }
        }
    }
}
