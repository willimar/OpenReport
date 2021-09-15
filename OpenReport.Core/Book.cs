using OpenReport.Core.Concrets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
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

        public List<ISheet> Sheets { get; protected set; }

        public async ValueTask<byte[]> BuildReport(int tasksCount, CancellationToken cancellationToken)
        {
            var parallelOptions = new ParallelOptions() 
            { 
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = tasksCount,
            };

            Parallel.ForEach(this.Sheets, parallelOptions, async sheet =>
            {
                foreach (var variable in sheet.Variables)
                {
                    this._bookProvider.set
                }
            });

            return await this._bookProvider.GetBuildedReport();
        }

        public async ValueTask SetValue(string variableName, string value, string sheetName)
        {
            var sheet = this.Sheets.FirstOrDefault(s => s.Name.ToLower() == sheetName.ToLower()) ?? throw new ArgumentNullException($"Sheet not found to {sheetName}");
            Func<IVariableField, bool> expression = (v) => v.FieldName.ToLower() == variableName;
            var variable = sheet.Variables.FirstOrDefault(expression) ?? throw new ArgumentNullException($"Variable not found to {variableName}");

            variable.Values.Add(new FieldValues(variable.Values.Count + 1, value));

            await ValueTask.CompletedTask;
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
            var sheets = this._bookProvider.GetSheetList();
            this.Sheets = new List<ISheet>();

            foreach (var sheet in sheets)
            {
                this.Sheets.Add(await this._bookProvider.GetSheetFrom(sheet));
            }
        }
    }
}
