using System.Linq;
using ExcelMapper.Tests.Models;
using Xunit;

namespace ExcelMapper.Tests
{
    public class Tests
    {
        [Fact]
        public void Test1()
        {
            var sut = new Excel(
                "sample.xlsx",
                p =>
                    p.Worksheet<WorksheetIsType>(q =>
                         q.Map(l => l.Header1)
                          .Map(l => l.Header2))
                     .Worksheet<NamedWorksheetWithHeaders>("Named worksheet", q =>
                         q.Map("Header1",l => l.Header1)
                          .Map("Header2",l => l.Header2)), 
                true);
            var act = sut.Get<NamedWorksheetWithHeaders>();
            Assert.Equal(3,act.Count());
        }
        
        [Fact]
        public void Test2()
        {
            var sut = new Excel(
                "sample.xlsx",
                p =>
                    p.Worksheet<WorksheetIsType>(q =>
                         q.Map(l => l.Header1)
                          .Map(l => l.Header2))
                     .Worksheet<NamedWorksheet>("Named worksheet", q =>
                         q.Map(1, l => l.Header1)
                          .Map(2, l => l.Header2)),
                true);
            var act = sut.Get<NamedWorksheet>();
            Assert.Equal(4,act.Count());
        }
        
        [Fact]
        public void Test3()
        {
            var sut = new Excel(
                "sample.xlsx",
                p =>
                    p.Worksheet<WorksheetIsType>(q =>
                         q.Map(l => l.Header1)
                          .Map(l => l.Header2)), 
                true);
            var act = sut.Get<WorksheetIsType>();
            Assert.Equal(3,act.Count());
        }
    }
}