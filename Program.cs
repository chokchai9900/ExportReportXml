using MongoDB.Driver;
using System;
using System.Linq;
using WebManageAPI.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace ExportReportXml
{
    class Program
    {
        static void Main(string[] args)
        {
            var partFile = @"D:\ReportExport.xlsx";
            var count = 1;
            Console.WriteLine("Export Report To Xml V1.0");
            Console.WriteLine($"File part = { partFile }");
            Console.WriteLine("please wait ......");

            var client = new MongoClient("mongodb://dbagent:Nso4Passw0rd5@mongodbproykgte5e7lvm7y-vm0.southeastasia.cloudapp.azure.com/nso");
            var database = client.GetDatabase("nso");
            var collectionReport = database.GetCollection<ReportEaInfo>("reporteainfo");
            var collectionSurvey = database.GetCollection<SurveyData>("survey");
            var countCells = 4;
            var getIDReport = collectionReport.Find(it=>true).ToList();
            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add($"EA Report");

                using (ExcelRange range = worksheet.Cells["A1:T3"])
                {
                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    range.Style.WrapText = true;
                    range.Style.Font.Bold = true;
                }
                using (ExcelRange range = worksheet.Cells["A3:T3"])
                {
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                }

                worksheet.Cells["H1:K1"].Merge = true;
                worksheet.Cells["H1"].Value = "สน1(อาคาร)";

                worksheet.Cells["L1:Q1"].Merge = true;
                worksheet.Cells["L1"].Value = "สน1(ครัวเรือน)";

                worksheet.Cells["R1:T1"].Merge = true;
                worksheet.Cells["R1"].Value = "สน2";

                worksheet.Cells["A1"].Value = "จังหวัด";
                worksheet.Cells["B1"].Value = "อำเภอ";
                worksheet.Cells["C1"].Value = "ตำบล";
                worksheet.Cells["D1"].Value = "เขตการปกครอง";
                worksheet.Cells["E1"].Value = "EA";
                worksheet.Cells["F1"].Value = "FS";
                worksheet.Cells["G1"].Value = "FI";
                worksheet.Cells["J2"].Value = "ทั้งหมด";
                worksheet.Cells["K2"].Value = "สมบูรณ์";

                worksheet.Cells["H2"].Value = "ทั้งหมด";
                worksheet.Cells["I2"].Value = "สมบูรณ์";
                worksheet.Cells["J2"].Value = "บ้านว่าง/ร้าง";
                worksheet.Cells["K2"].Value = "3ครั้งไม่พบ";

                worksheet.Cells["L2"].Value = "ทั้งหมด";
                worksheet.Cells["M2"].Value = "สมบูรณ์";
                worksheet.Cells["N2"].Value = "บ้านว่าง/ร้าง";
                worksheet.Cells["O2"].Value = "3ครั้งไม่พบ";
                worksheet.Cells["P2"].Value = "ยุติสมบูรณ์";
                worksheet.Cells["Q2"].Value = "ยุติไม่สมบูรณ์";

                worksheet.Cells["R2"].Value = "ทั้งหมด";
                worksheet.Cells["S2"].Value = "สมบูรณ์";
                worksheet.Cells["T2"].Value = "ไม่สมบูรณ์";

                    foreach (var item in getIDReport)
                    {
                    var getdataSurvey = collectionSurvey.Find(x => x.EA == item._id && x.DeletionDateTime == null && x.Enlisted == true ).ToList();

                    if (getdataSurvey == null)
                    {

                        worksheet.Cells[countCells, 1].Value = item.REG_NAME;
                        worksheet.Cells[countCells, 2].Value = item.AMP_NAME;
                        worksheet.Cells[countCells, 3].Value = item.TAM_NAME;
                        worksheet.Cells[countCells, 4].Value = item.DISTRICT;
                        worksheet.Cells[countCells, 5].Value = item.EA;
                        worksheet.Cells[countCells, 6].Value = item.ApproveByFs == null ? "0" : item.ApproveByFs;
                        worksheet.Cells[countCells, 7].Value = "-";

                        worksheet.Cells[countCells, 8].Value = "-";
                        worksheet.Cells[countCells, 9].Value = "-";
                        worksheet.Cells[countCells, 10].Value = "-";
                        worksheet.Cells[countCells, 11].Value = "-";

                        worksheet.Cells[countCells, 12].Value = "-";
                        worksheet.Cells[countCells, 13].Value = "-";
                        worksheet.Cells[countCells, 14].Value = "-";
                        worksheet.Cells[countCells, 15].Value = "-";
                        worksheet.Cells[countCells, 16].Value = "-";
                        worksheet.Cells[countCells, 17].Value = "-";

                        worksheet.Cells[countCells, 18].Value = "-";
                        worksheet.Cells[countCells, 19].Value = "-";
                        worksheet.Cells[countCells, 20].Value = "-";

                    }
                    else
                    {
                        var GroupDataByID = getdataSurvey
                            .GroupBy(it => it.UserId)
                            .ToList();
                        foreach (var list in GroupDataByID)
                        {
                            //building
                            var AllCountBuild = list.Count(it => it.SampleType == "b");
                            var CompletCountBuild = list.Count(it => it.Status == "done-all" && it.SampleType == "b");
                            var sadCountBuild = list.Count(it => it.Status == "sad" && it.SampleType == "b");
                            var eye_offCountBuild = list.Count(it => it.Status == "eye-off" && it.SampleType == "b");

                            var AllCountUnit = list.Count(it => it.SampleType == "u");
                            var CompletCountUnit = list.Count(it => it.Status == "complete" && it.SampleType == "u");
                            var sadCountUnit = list.Count(it => it.Status == "sad" && it.SampleType == "u");
                            var eye_offCountUnit = list.Count(it => it.Status == "eye-off" && it.SampleType == "u");
                            var mic_offCountUnit = list.Count(it => it.Status == "mic-off" && it.SampleType == "u");
                            var pauseCountUnit = list.Count(it => it.Status == "pause" || it.Status == "refresh" && it.SampleType == "u");

                            var commUnityTypeAll = list.Count(it => it.SampleType == "c");
                            var CompletcommUnityType = list.Count(it => it.SampleType == "c" && it.Status == "done-all");
                            var unCompletcommUnityType = list.Count(it => it.SampleType == "c" && it.Status != "done-all");

                            worksheet.Cells[countCells, 1].Value = item.CWT_NAME;
                            worksheet.Cells[countCells, 2].Value = item.AMP_NAME;
                            worksheet.Cells[countCells, 3].Value = item.TAM_NAME;
                            worksheet.Cells[countCells, 4].Value = item.DISTRICT;
                            worksheet.Cells[countCells, 5].Value = item.EA;
                            worksheet.Cells[countCells, 6].Value = item.ApproveByFs == null ? "0" : item.ApproveByFs;
                            worksheet.Cells[countCells, 7].Value = list.Key;

                            worksheet.Cells[countCells, 8].Value = AllCountBuild;
                            worksheet.Cells[countCells, 9].Value = CompletCountBuild;
                            worksheet.Cells[countCells, 10].Value = eye_offCountBuild;
                            worksheet.Cells[countCells, 11].Value = sadCountBuild;

                            worksheet.Cells[countCells, 12].Value = AllCountUnit;
                            worksheet.Cells[countCells, 13].Value = CompletCountUnit;
                            worksheet.Cells[countCells, 14].Value = eye_offCountUnit;
                            worksheet.Cells[countCells, 15].Value = sadCountUnit;
                            worksheet.Cells[countCells, 16].Value = mic_offCountUnit;
                            worksheet.Cells[countCells, 17].Value = pauseCountUnit;

                            worksheet.Cells[countCells, 18].Value = commUnityTypeAll;
                            worksheet.Cells[countCells, 19].Value = CompletcommUnityType;
                            worksheet.Cells[countCells, 20].Value = unCompletcommUnityType;

                            Console.Write($"{count} : ");
                            Console.WriteLine($"{list.Key} Complet !!");
                            count++;
                            countCells++;
                        }
                    }
                }
                excelPackage.SaveAs(new FileInfo(partFile));
            }
        }
    }
}
