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
            const string filePath = @"C:\Users\peboos\Documents\ABG Gavlegården\arbetsordertyper.xlsx";
            const string resultPath = @"C:\Users\peboos\Documents\ABG Gavlegården\arbetsordertyper_json.txt";
            const string sheetName = "Sammanställning";

            var excelRows = new List<ExcelRow>();
            var list = new List<Location>();

            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                var s = p.Workbook.Worksheets.FirstOrDefault(n => n.Name == sheetName);
                if (s == null)
                {
                    Console.WriteLine($"Sheet 'sheetName' not found.");
                    return false;
                }

                var rows = s.Dimension.End.Row;

                for (var row = 3; row <= rows; row++)
                {
                    excelRows.Add(new ExcelRow
                    {
                        Location = s.Cells[row, 1].Value?.ToString(),
                        SpaceCode = s.Cells[row, 2].Value?.ToString(),
                        SpaceCaption = s.Cells[row, 3].Value?.ToString(),
                        PartCode = s.Cells[row, 4].Value?.ToString(),
                        PartCaption = s.Cells[row, 5].Value?.ToString(),
                        WorkOrderType = s.Cells[row, 6].Value?.ToString(),
                        WorkOrderCategory = s.Cells[row, 7].Value?.ToString()
                    });
                }
            }

            // Locations
            foreach (var row in excelRows)
            {
                if (list.Any(l => l.Caption == row.Location)) continue;
                list.Add(new Location { Caption = row.Location, Spaces = new List<Space>() });
            }

            // Spaces
            foreach (var row in excelRows)
            {
                var location = list.SingleOrDefault(l => l.Caption == row.Location);
                if (location.Spaces.Any(x => x.Caption == row.SpaceCaption && x.Code == row.SpaceCode)) continue;
                location.Spaces.Add(new Space { Caption = row.SpaceCaption, Code = row.SpaceCode, BuildingParts = new List<Part>() });
            }

            // Parts
            foreach (var row in excelRows)
            {
                var location = list.SingleOrDefault(l => l.Caption == row.Location);
                var space = location.Spaces.SingleOrDefault(x => x.Caption == row.SpaceCaption && x.Code == row.SpaceCode);
                if (space.BuildingParts.Any(x =>
                    x.Caption == row.PartCaption &&
                    x.Code == row.PartCode &&
                    x.WorkOrder.Type == row.WorkOrderType &&
                    x.WorkOrder.Category == row.WorkOrderCategory)) continue;
                space.BuildingParts.Add(new Part
                {
                    Caption = row.PartCaption,
                    Code = row.PartCode,
                    WorkOrder = new WorkOrder
                    {
                        Type = row.WorkOrderType,
                        Category = row.WorkOrderCategory
                    }
                });
            }

            var json = JsonConvert.SerializeObject(list, Formatting.Indented);

            using (var sw = new StreamWriter(resultPath))
            {
                sw.Write(json);
            }

            return true;
        }
    }
}
