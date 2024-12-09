using NUnit.Framework;

namespace TestProject1
{
    public class CellReferenceTests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        [TestCase(1u, 1u, "A1")]
        [TestCase(26u, 1u, "Z1")]
        [TestCase(27u, 1u, "AA1")]
        [TestCase(28u, 1u, "AB1")]
        [TestCase(26u * 2u, 1u, "AZ1")]
        [TestCase(26u * 2u + 1, 1u, "BA1")]
        [TestCase(26u * 3u, 1u, "BZ1")]
        [TestCase(26u * 27u, 10u, "ZZ10")]
        [TestCase(901u, 10u, "AHQ10")]
        public void Test1(uint column, uint row, string text)
        {
            {
                var pos = new TSKT.Bonn.CellReference(column, row);
                Assert.Equals(text, pos.value);
            }
            {
                var pos = new TSKT.Bonn.CellReference(text);
                Assert.Equals(column, pos.columnIndex);
                Assert.Equals(row, pos.rowIndex);
            }
        }

        [Test]
        public void Test2()
        {
            for (uint i = 1; i < 32768; ++i)
            {
                var pos = new TSKT.Bonn.CellReference(i, i);
                var b = new TSKT.Bonn.CellReference(pos.value);
                Assert.Equals(i, b.columnIndex);
                Assert.Equals(i, b.rowIndex);
            }
        }
    }
}