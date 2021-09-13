using FluentAssertions;
using Moq.AutoMock;
using OpenReport.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OpenReport.Tests.UnitTests.BookTests
{
    public class BookTest
    {
        private readonly AutoMocker _mocker;

        public BookTest()
        {
            this._mocker = new AutoMocker();
        }

        [Fact]
        public async Task Book_Contructor_Test()
        {
            await using var book = this._mocker.CreateInstance<Book>();

            book.Should().NotBeNull();
        }

    }
}
