using Moq.AutoMock;
using OpenReport.Core;
using OpenReport.Tests.Properties;
using OpenXml.Provider;
using System.Collections.Generic;
using System.IO;
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

        [Theory]
        [InlineData("Willimar Augusto Rocha", 44, "08577952684", "Poly Collor LTDA", 
            new string[] { "Willimar Augusto Rocha", "Fernanda S S Rocha", "Thiago S Rocha" },
            new string[] { "Tech Lead", "Hair Designer", "" },
            new string[] { "1500.85", "1854.66", "333.44" })]
        public async Task Criar_Modelo03(string name, int yearsOld, string document, string companyName, string[] employears, string[] functions, string[] salaries)
        {
            var modelo01 = new MemoryStream(Resources.Modelo03);
            await using var bookProvider = new BookProvider(modelo01);

            await using var book = new Book(bookProvider);

            await book.SetValue("nome", name, "Sheet2");
            await book.SetValue("idade", yearsOld.ToString(), "Sheet2");
            await book.SetValue("documento", document, "Sheet2");
            await book.SetValue("razao.social", companyName, "Sheet2");

            foreach (var item in employears)
            {
                await book.SetValue("funcionario", item, "Sheet2");
            }

            foreach (var item in functions)
            {
                await book.SetValue("cargo", item, "Sheet2");
            }

            foreach (var item in salaries)
            {
                await book.SetValue("salario", item, "Sheet2");
            }

            await book.BuildReport(5, new CancellationToken());
        }
    }
}
