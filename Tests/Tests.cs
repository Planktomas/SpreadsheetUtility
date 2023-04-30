using SpreadsheetUtility;

namespace Tests
{
    [TestFixture]
    public class SpreadsheetTests
    {
        [TestCase(0, 0, "A1", TestName = "Cell: 0x0 -> A1")]
        [TestCase(6, 1, "G2", TestName = "Cell: 6x1 -> G2")]
        [TestCase(25, 0, "Z1", TestName = "Cell: 25x0 -> Z1")]
        [TestCase(26, 0, "AA1", TestName = "Cell: 26x0 -> AA1")]
        [TestCase(27, 0, "AB1", TestName = "Cell: 27x0 -> AB1")]
        [TestCase(30, 0, "AE1", TestName = "Cell: 30x0 -> AE1")]
        [TestCase(51, 0, "AZ1", TestName = "Cell: 51x0 -> AZ1")]
        [TestCase(52, 0, "BA1", TestName = "Cell: 52x0 -> BA1")]
        [TestCase(701, 0, "ZZ1", TestName = "Cell: 701x0 -> ZZ1")]
        [TestCase(702, 0, "AAA1", TestName = "Cell: 702x0 -> AAA1")]
        [TestCase(950, 950, "AJO951", TestName = "Cell: 950x950 -> AJO951")]
        [TestCase(2023, 2023, "BYV2024", TestName = "Cell: 2023x2023 -> BYV2024")]
        public void Cell_ConvertsXYToCellCoordinates(int x, int y, string expectedResult)
        {
            Assert.AreEqual(expectedResult, Spreadsheet.Cell(x, y));
        }
    }
}