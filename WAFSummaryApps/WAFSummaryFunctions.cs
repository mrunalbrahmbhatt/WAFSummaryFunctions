using CsvHelper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
            Dictionary<string, PillarInfo> pillarInfo = null;

            var categoryDescriptions = GetCategoryDescription(csvDescription);

            string fileContent = GetContent(csvWAF);
            var wafRecords = GetWafRecords(fileContent);
            var pillars = categoryDescriptions.Select(r => r.Pillar).Distinct().ToList();

            //Init Pillar Info
            pillarInfo = InitPillarInfo(pillars);


            string overallScore = string.Empty;
            GetPillarScore(pillarInfo, fileContent, out overallScore);
            GetPillarDescription(pillarInfo, categoryDescriptions,pillars);

            var outputBlobName = $"WAFSummary{DateTime.Today:ddMMyyyy}.pptx";
            var outboundBlob = new BlobAttribute($"powerpoint/{outputBlobName}", FileAccess.Write);
            using (var ppt = new StreamWriter(pptBinder.Bind<Stream>(outboundBlob)))
            {
                await ppt.WriteLineAsync("hello");
            }
            log.LogInformation("C# HTTP trigger function processed a request.");
            //string requestBody = await new StreamReader(csv.Content).ReadToEndAsync();
            //dynamic data = JsonConvert.SerializeObject(csvContent.Path);
            return new OkObjectResult($"powerpoint/{outputBlobName}");
        }

        

        private void GetPillarDescription(Dictionary<string, PillarInfo> pillarInfo, IEnumerable<WafCategoryDescription> categoryDescriptions, List<string> pillars)
        {
            foreach(var pillar in pillars)
            {
                pillarInfo[pillar].Description = categoryDescriptions
                  .Where(d=> d.Pillar.Contains(pillar) && d.Category.Equals("Survey Level Group")).Select(d=>d.Description).FirstOrDefault();
            }
        }

        private Dictionary<string, PillarInfo> InitPillarInfo(List<string> pillars)
        {
            Dictionary<string, PillarInfo> pillarInfo  = new Dictionary<string, PillarInfo>();
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

        //private IEnumerable<WAFRecord> GetWafRecords(Stream csvWAF)
        //{
        //    using var r = new StreamReader(csvWAF,System.Text.Encoding.ASCII);
        //    var csvContent = r.ReadToEnd();
        //    int start = csvContent.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context");
        //    int end = csvContent.IndexOf("-----------,,,,,");
        //    string test = csvContent[start..end];
        //    using var reader = new StringReader(test);
        //    using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        //    return csv.GetRecords<WAFRecord>();
        //}
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
