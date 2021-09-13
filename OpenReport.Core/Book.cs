using OpenReport.Core.Concrets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OpenReport.Core
{
    public class Book : IAsyncDisposable
    {
        private readonly IBookProvider _bookProvider;

        public Book(IBookProvider bookProvider)
        {
            this._bookProvider = bookProvider ?? throw new ArgumentNullException(nameof(bookProvider));
            _ = this.OnConstructor();
        }

        private Book() 
        {
            throw new InvalidOperationException();
        }

        public IEnumerable<ISheet> Sheets { get; protected set; }

        public async ValueTask<byte[]> BuildReport(int maxDegreeOfParallelism, CancellationToken cancellationToken)
        {
            var parallelOptions = new ParallelOptions() 
            { 
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = maxDegreeOfParallelism,
            };

            Parallel.ForEach(this.Sheets, parallelOptions, async sheet => 
            {
                var sheetBuilder = await this._bookProvider.SheetBuilderCreate(sheet.Name);
                var fieldList = sheet.Fields.OrderBy(sht => sht.RowPos);

                foreach (var field in fieldList)
                {
                    if (field.FieldType.Equals(FieldType.repeatGroup))
                    {
                        var values = field.FieldValue as IRepeatGroup;
                        uint linesToInsert = await this.GetCountLinesToInsert(values);

                        await this.InsertLinesBellow(values.EndRow, linesToInsert, sheet);

                        await this.SetValue(values, sheet);
                    }
                    else
                    {
                        await this.SetValue(field, sheet);
                    }
                }
            });

            return await this._bookProvider.GetBuildedReport();
        }

        private Task SetValue(IField field, ISheet sheet)
        {
            throw new NotImplementedException();
        }

        private Task SetValue(IRepeatGroup values, ISheet sheet)
        {
            throw new NotImplementedException();
        }

        private Task InsertLinesBellow(uint endRow, uint linesToInsert, ISheet sheet)
        {
            throw new NotImplementedException();
        }

        private Task<uint> GetCountLinesToInsert(IRepeatGroup values)
        {
            throw new NotImplementedException();
        }

        public async ValueTask DisposeAsync()
        {
            await this.DisposeAsync(true);
            GC.SuppressFinalize(this);
        }

        protected virtual async ValueTask DisposeAsync(bool disposing)
        {
            if (disposing)
            {
                
            }

            await ValueTask.CompletedTask;
        }

        protected virtual async ValueTask OnConstructor()
        {
            string[] sheetNameList = this._bookProvider.GetSheetList();

            foreach (var sheetName in sheetNameList)
            {
                IEnumerable<(string cellAddress, string variableName)> variableList = this._bookProvider.GetVariableList(sheetName);

                this.Sheets.Append(new Sheet() { 
                    Name = sheetName,
                    Fields = await this.GetFields(variableList, sheetName)
                });
            }
        }

        protected async ValueTask<IEnumerable<IField>> GetFields(IEnumerable<(string cellAddress, string variableName)> variableList, string sheetName)
        {
            var fields = new List<IField>();

            foreach (var cellInfo in variableList)
            {
                IField field = new Field(cellInfo.cellAddress, cellInfo.variableName)
                {
                    FieldValue = await this._bookProvider.GetFieldValueTo(cellInfo.cellAddress, cellInfo.variableName, sheetName),
                    Style = await this._bookProvider.GetFieldStyleTo(cellInfo.cellAddress, cellInfo.variableName, sheetName),
                    Formula = await this._bookProvider.GetFormulaTo(cellInfo.cellAddress, cellInfo.variableName, sheetName),
                    FieldType = await this._bookProvider.GetFieldType(cellInfo.cellAddress, cellInfo.variableName, sheetName),
                };

                if (field.FieldType.Equals(FieldType.repeatGroup) && !field.FieldValue.GetType().Equals(typeof(IRepeatGroup)))
                {
                    throw new InvalidOperationException($"Please use a IRepeatGroup to FieldValue when use FieldType equals repeatGroup.");
                }

                fields.Add(field);
            }

            return await ValueTask.FromResult(fields);
        }
    }
}
