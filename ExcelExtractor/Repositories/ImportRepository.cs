using ExcelExtractor.Attributes;
using ExcelExtractor.Contracts;
using ExcelExtractor.Models.ExcelModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Reflection;

namespace ExcelExtractor.Repositories
{
    public class ImportRepository
    {
        private readonly Serilog.ILogger _logger;

        public ImportRepository()
        {
            _logger = Serilog.Log.Logger;
        }

        public string ImportExcelFile(IFormFile file)
        {
            var importPath = "Uploads";
            var dirPath = Path.Combine(Directory.GetCurrentDirectory(), importPath + "/", DateTime.Now.ToString("yyyyMMdd_HHmmsss"));
            try {
                if (!Directory.Exists(dirPath)) {
                    Directory.CreateDirectory(dirPath);
                }

                var fileName = file.FileName;
                var fileExt = fileName.Split('.').Last();
                var filePath = Path.Combine(dirPath, fileName);
                IWorkbook workbook;
                using (var fileStream = new FileStream(filePath, FileMode.Create)) {
                    file.CopyTo(fileStream);

                    fileStream.Position = 0;
                    if (fileExt == "xlsx") {
                        workbook = new XSSFWorkbook(fileStream);
                    }
                    else {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                }

                ReadWorkbook(workbook);
            }
            catch (Exception e) {
                Serilog.Log.Logger.Error(e, e.Message);
            }
            finally {
                if (Directory.Exists(dirPath)) {
                    Directory.Delete(dirPath, true);
                }
            }
        }

        public IEnumerable<IExcelModel> ReadWorkbook(IWorkbook workbook)
        {
            var sheetIndex = 0;
            var sheet = workbook.GetSheetAt(sheetIndex);

            // GetHeaders
            var headers = new List<string>();
            var headerRow = sheet.GetRow(0);
            for (var cellIndex = 0; cellIndex <= headerRow.LastCellNum; cellIndex++) {
                var headerCell = headerRow.GetCell(cellIndex);
                if (headerCell == null || string.IsNullOrEmpty(headerCell.StringCellValue)) continue;

                headers.Add(headerCell.StringCellValue.Trim());
            }

            var valueList = new List<Dictionary<string, object?>>();
            for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++) {
                var skipRow = false;
                var row = sheet.GetRow(rowIndex);
                if (row == null) break;

                var values = new Dictionary<string, object?>();
                for (var cellIndex = 0; cellIndex < headers.Count; cellIndex++) {
                    var cell = row.GetCell(cellIndex);
                    var propName = headers.ElementAt(cellIndex);

                    if (cell == null || string.IsNullOrEmpty(cell.ToString())) {
                        if (cellIndex == 0) skipRow = true;
                        continue;
                    }

                    var cellValue = GetCellValue(cell);
                    values.Add(propName, cellValue);
                }

                if (skipRow) continue;
                valueList.Add(values);
            }

            var jHeaders = JArray.FromObject(headers);
            var jValues = JArray.FromObject(valueList);
            var results = ConvertDictionaryToModel(jHeaders, jValues);

            return results;
        }

        private IEnumerable<IExcelModel> ConvertDictionaryToModel(JArray headers, JArray values)
        {
            if (GetModelByHeaders(headers) == typeof(MemberExcel)) {

            }
        }

        private Type GetModelByHeaders(JArray headers)
        {
            var modelTypes = new[] {
                typeof(MemberExcel),
            };
            foreach(var type in modelTypes) {
                var props = type.GetProperties(BindingFlags.Public)
                    .Where(x => x.GetCustomAttribute<ColKeyAttribute>() != null)
                    .Select(x => new {
                        Key = x.GetCustomAttribute<ColKeyAttribute>()?.Key ?? string.Empty,
                        
                    });
            }
        }

        private object? GetCellValue(ICell cell, CellType? cellType = null)
        {
            cellType ??= cell.CellType;

            switch (cellType) {
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                default:
                    string? val = cell.ToString();
                    if (string.IsNullOrEmpty(val)) {
                        val = null;
                    }
                    return val;
            }
        }
    }
}
