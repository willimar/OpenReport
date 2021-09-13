using Moq.AutoMock;
using Npoi.Provider;
using OpenReport.Core;
using OpenReport.Tests.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OpenReport.Tests.IntegrationTests.BookTests
{
    public class BookTest
    {
        private readonly AutoMocker _mocker;

        public BookTest()
        {
            this._mocker = new AutoMocker();
        }

        [Fact]
        public async Task Criar_Modelo03()
        {
            var modelo01 = new MemoryStream(Resources.Modelo03);
            await using var bookProvider = new BookProvider(modelo01);

            await using var book = new Book(bookProvider);

            await book.BuildReport(5, new CancellationToken());
        }
    }
}
