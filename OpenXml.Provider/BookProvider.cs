using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenReport.Core;
using OpenReport.Core.Concrets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Provider
{
    public class BookProvider : IBookProvider, IAsyncDisposable
    {
        private SpreadsheetDocument _spreadsheetDocument;
        private MemoryStream _streamBook;

        public BookProvider(Stream excelBook)
        {
            _ = this.OnContructor(excelBook);
        }

        private BookProvider()
        {

        }

        protected virtual async ValueTask DisposeAsync(bool disposing)
        {
            if (disposing)
            {
                this._spreadsheetDocument.Dispose();
                await this._streamBook.DisposeAsync();                
            }

            await ValueTask.CompletedTask;
        }

        protected virtual async ValueTask OnContructor(Stream excelBook)
        {
            try
            {
                this._streamBook = new MemoryStream();
                excelBook.CopyTo(this._streamBook);
                this._spreadsheetDocument = SpreadsheetDocument.Open(this._streamBook, true);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                excelBook.Close();
                await excelBook.DisposeAsync();
            }
        }

        public async ValueTask DisposeAsync()
        {
            await this.DisposeAsync(true);
        }

        public ValueTask<byte[]> GetBuildedReport()
        {
            throw new NotImplementedException();
        }

        public ValueTask<IStyle> GetFieldStyleTo(string cellAddress, string variableName, string sheetName)
        {
            throw new NotImplementedException();
        }

        public ValueTask<FieldType> GetFieldType(string cellAddress, string variableName, string sheetName)
        {
            throw new NotImplementedException();
        }

        public ValueTask<string> GetFieldValueTo(string cellAddress, string variableName, string sheetName)
        {
            throw new NotImplementedException();
        }

        public ValueTask<object> GetFormulaTo(string cellAddress, string variableName, string sheetName)
        {
            throw new NotImplementedException();
        }

        public string[] GetSheetList()
        {
            var workbookPart = this._spreadsheetDocument.WorkbookPart;
            var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.Name.ToString());

            return sheets.ToArray();
        }

        public async ValueTask<IEnumerable<IField>> GetVariableList(string sheetName)
        {
            var result = new List<IField>();
            var sheet = await this.GetSheetFrom(sheetName);

            foreach (var row in sheet.Rows.Where(r => r.Fields.Any(f => f.IsVariable)))
            {
                result.AddRange(row.Fields.Where(f => f.IsVariable));
            }

            return await ValueTask.FromResult(result);
        }

        public ValueTask<ISheetBuilder> SheetBuilderCreate(string name)
        {
            throw new NotImplementedException();
        }

        public async ValueTask<ISheet> GetSheetFrom(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var workbookPart = this._spreadsheetDocument.WorkbookPart;
            var workSheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(s => s.Name.Equals(sheetName)).FirstOrDefault();
            
            if (workSheet is null)
            {
                throw new ArgumentException($"Don't have a sheet to the name '{sheetName}' in the work book.");
            }

            var workSheetPart = (WorksheetPart)workbookPart.GetPartById(workSheet.Id);

            var sheet = new OpenSheet(sheetName);

            sheet.Rows.AddRange(this.GetRowsFrom(workSheetPart, workbookPart));

            return await ValueTask.FromResult(sheet);
        }

        private IEnumerable<IRow> GetRowsFrom(WorksheetPart workSheet, WorkbookPart workbookPart)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = workSheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>();
            var result = new List<OpenRow>();;

            foreach (var row in rows)
            {
                var typedRow = new OpenRow(Convert.ToInt32(row.RowIndex.Value));

                typedRow.Fields.AddRange(this.GetFieldsFrom(row, workbookPart, workSheet));

                result.Add(typedRow);
            }

            return result;
        }

        private IEnumerable<IField> GetFieldsFrom(Row row, WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            var fields = new List<IField>();
            var cells = row.Descendants<Cell>();
            
            foreach (var cell in cells)
            {
                var field = new OpenField(cell.CellReference.Value, this.GetValue(cell, workbookPart));
                var mergedCells = worksheetPart.Worksheet.Descendants<MergeCell>();
                var mergedAddress = $"{field.CellAddress}:";
                var isMerged = mergedCells.Any(m => m.Reference.Value.Contains(mergedAddress));
                var mergedCellsAddress = isMerged ? mergedCells.Where(m => m.Reference.Value.Contains(mergedAddress)).First().Reference.Value : string.Empty;

                field.Style = new OpenStyle()
                {
                    StyleIndex = Convert.ToInt32(cell.StyleIndex?.Value),
                    IsMerged = isMerged,
                    MergeCells = mergedCellsAddress,
                };

                fields.Add(field);                
            }

            return fields;
        }

        private string GetValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell is null || cell.DataType == null)
            {
                return string.Empty;
            }

            if (cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (stringTable != null)
                {
                    var value = stringTable.SharedStringTable.ElementAt(int.Parse(cell.InnerText, null)).InnerText;

                    return value;
                }
            }

            return cell.InnerText;
        }
    }
}
