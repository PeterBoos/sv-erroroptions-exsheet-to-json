using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public static class TranslateExcelSheet
    {
        public static bool Run()
        {
            const string filePath = @"C:\Users\peboos\Documents\ABG Gavlegårdarna\Etapp 3\Felrapport\Arbetsorderreferens\arbetsordertyper.xlsx";
            const string resultPath = @"C:\Users\peboos\Documents\ABG Gavlegårdarna\Etapp 3\Felrapport\Arbetsorderreferens\arbetsordertyper_json_20190523.txt";
            const string sheetName = "Sammanställning";

            var excelRows = new List<ExcelRow>();
            var options = new Options { Areas = new List<Area>() };
            var list = options.Areas;

            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                var s = p.Workbook.Worksheets.FirstOrDefault(n => n.Name == sheetName);
                if (s == null)
                {
                    Console.WriteLine($"Sheet '{sheetName}' not found.");
                    return false;
                }

                var rows = s.Dimension.End.Row;

                for (var row = 3; row <= rows; row++)
                {
                    excelRows.Add(new ExcelRow
                    {
                        AreaCaption = s.Cells[row, 1].Value?.ToString(),
                        AreaCode = s.Cells[row, 2].Value?.ToString(),
                        LocationCode = s.Cells[row, 3].Value?.ToString(),
                        LocationCaption = s.Cells[row, 4].Value?.ToString(),
                        PartCode = s.Cells[row, 5].Value?.ToString(),
                        PartCaption = s.Cells[row, 6].Value?.ToString(),
                        WorkOrderType = s.Cells[row, 7].Value?.ToString(),
                        WorkOrderCategory = s.Cells[row, 8].Value?.ToString()
                    });
                }
            }

            // Areas
            foreach (var row in excelRows)
            {
                if (list.Any((Func<Area, bool>)(a => a.Code == row.AreaCode))) continue;
                list.Add(new Area { Caption = row.AreaCaption, Code = row.AreaCode, Locations = new List<Location>() });
            }

            // Locations
            foreach (var row in excelRows)
            {
                var area = list.SingleOrDefault((Func<Area, bool>)(l => l.Code == row.AreaCode));
                if (area.Locations.Any(x => x.Code == row.LocationCode && x.Caption == row.LocationCaption)) continue;
                area.Locations.Add(new Location { Caption = row.LocationCaption, Code = row.LocationCode, BuildingParts = new List<BuildingPart>() });
            }

            // Parts
            foreach (var row in excelRows)
            {
                var area = list.SingleOrDefault((Func<Area, bool>)(l => l.Code == row.AreaCode));
                var location = area.Locations.SingleOrDefault(x => x.Caption == row.LocationCaption && x.Code == row.LocationCode);
                if (location.BuildingParts.Any(x =>
                    x.Caption == row.PartCaption &&
                    x.Code == row.PartCode)) continue;
                    //x.WorkOrder.Type == row.WorkOrderType &&
                    //x.WorkOrder.Category == row.WorkOrderCategory)) continue;
                location.BuildingParts.Add(new BuildingPart
                {
                    Caption = row.PartCaption,
                    Code = row.PartCode
                    //WorkOrder = new WorkOrder
                    //{
                    //    Type = row.WorkOrderType,
                    //    Category = row.WorkOrderCategory
                    //}
                });
            }

            var json = JsonConvert.SerializeObject(options, Formatting.Indented);

            using (var sw = new StreamWriter(resultPath))
            {
                sw.Write(json);
            }

            return true;
        }
    }
}
