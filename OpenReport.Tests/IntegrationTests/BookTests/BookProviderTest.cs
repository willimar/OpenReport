using FluentAssertions;
using OpenReport.Tests.Properties;
using OpenXml.Provider;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OpenReport.Tests.IntegrationTests.BookTests
{
    public class BookProviderTest
    {
        [Fact]
        public async Task SheetsNameCheckTest()
        {
            var modelo = new MemoryStream(Resources.Modelo03);
            await using var bookProvider = new BookProvider(modelo);

            var sheets = bookProvider.GetSheetList();

            sheets.Should().NotBeNull();
            sheets.Should().HaveCount(3);
            sheets.Should().Contain(new string[] { "Sheet2", "Sheet3", "Sheet4" });
        }

        [Theory]
        [InlineData("A1", "MODELO PARA TESTES DE EXECUÇÃO", true, 1)]
        [InlineData("A3", "Informações contidas", true, 3)]
        [InlineData("B4", "Idade", false, 4)]
        [InlineData("C10", "${salario}", false, 10)]
        [InlineData("A13", "A IDEIA É EMPURRAR O RODAPÉ", true, 13)]
        public async Task SheetObjectCheckTest(string fieldAddress, string fiedValue, bool isMerged, int rowIndex)
        {
            var modelo = new MemoryStream(Resources.Modelo03);
            await using var bookProvider = new BookProvider(modelo);

            var sheets = await bookProvider.GetSheetFrom("Sheet2");

            sheets.Should().NotBeNull();

            sheets.Rows.Should().ContainSingle(r => r.RowIndex == rowIndex && r.Fields.Any(f => f.CellAddress == fieldAddress && f.Style.IsMerged == isMerged && f.FieldValue == fiedValue));
            sheets.Rows.Should().HaveCount(11);
            sheets.Rows.Where(x => x.Fields.Count == 3).Should().HaveCount(11);
        }
    }
}
