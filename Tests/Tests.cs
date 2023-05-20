using SpreadsheetUtility;

namespace Tests
{
    [TestFixture]
    public class SpreadsheetTests
    {
        const string k_SpreadsheetFilename = "test.xlsx";
        const string k_SpreadsheetFilename2 = "test2.xlsx";
        const string k_FolderName = "Folder/";

        class Simple
        {
            public string Text { get; set; }

            public Simple()
            {
                Text = Guid.NewGuid().ToString();
            }

            public Simple(string text)
            {
                Text = text;
            }

            public static Simple[] One = new[] { new Simple() };
            public static Simple[] Three = new[] { new Simple("1"), new Simple("2"), new Simple("3") };
        }

        class AutoFit
        {
            public string? S { get; set; }
            public string? LongLongLongLong { get; set; }

            public AutoFit(string? @short, string? @long)
            {
                S = @short;
                LongLongLongLong = @long;
            }

            public AutoFit() : this(null, null) { }

            public static AutoFit[] Headers = new[] { new AutoFit() };
            public static AutoFit[] Values = new[] { new AutoFit("Short", "LongLongLongLongLongLong") };
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(k_FolderName))
                Directory.Delete(k_FolderName, true);

            File.Delete(k_SpreadsheetFilename);
            File.Delete(k_SpreadsheetFilename2);
        }

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
            Assert.That(Spreadsheet.Cell(x, y), Is.EqualTo(expectedResult));
        }

        [Test]
        public void Constructor_OpensOrCreatesSpreadsheet()
        {
            Assert.That(File.Exists(k_SpreadsheetFilename), Is.False);

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                Assert.That(spreadsheet?.document.GetCellValueAsString("A1"), Is.Empty);

                spreadsheet.Write(Simple.One);

                Assert.That(spreadsheet?.document.GetCellValueAsString("A1"), Is.Not.Empty);
            }

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                Assert.That(spreadsheet?.document.GetCellValueAsString("A1"), Is.Not.Empty,
                    "When reopening spreadsheet we should be able to read data that is in it");
            }
        }

        [Test]
        public void Destructor_WhenSavingSpreadsheetInNonExistingFolder_Throws()
        {
            Assert.Throws<DirectoryNotFoundException>(() =>
            {
                using var spreadsheet = new Spreadsheet(Path.Combine(k_FolderName, k_SpreadsheetFilename));
                spreadsheet.Write(Simple.One);
            });
        }

        [Test]
        public void Destructor_WhenSavingSpreadsheetInFolder_DoesNotThrow()
        {
            Directory.CreateDirectory(k_FolderName);

            Assert.DoesNotThrow(() =>
            {
                using var spreadsheet = new Spreadsheet(Path.Combine(k_FolderName, k_SpreadsheetFilename));
                spreadsheet.Write(Simple.One);
            });
        }

        [Test]
        public void WritingDisposedSpreadsheet_Throws()
        {
            var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Dispose();

            Assert.Throws<ObjectDisposedException>(() => spreadsheet.Write(Simple.One));
        }

        [Test]
        public void DisposingAlreadyDisposedSpreadsheet_Throws()
        {
            Spreadsheet spreadsheet = new(k_SpreadsheetFilename);
            spreadsheet.Dispose();

            Assert.Throws<ObjectDisposedException>(() => spreadsheet.Dispose());
        }

        [Test]
        public void Headers_InfluencesColumnWidth()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(AutoFit.Headers);

            var shortColumnPrevWidth = spreadsheet?.document.GetColumnWidth("A1");
            var longColumnPrevWidth = spreadsheet?.document.GetColumnWidth("B1");

            spreadsheet?.AutoFit();

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet?.document.GetColumnWidth("A1"), Is.LessThan(shortColumnPrevWidth));
                Assert.That(spreadsheet?.document.GetColumnWidth("B1"), Is.GreaterThan(longColumnPrevWidth));
            });
        }

        [Test]
        public void Values_InfluencesColumnWidth()
        {
            double? shortHeaderWidth;
            double? longHeaderWidth;

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                spreadsheet.Write(AutoFit.Headers);
                spreadsheet.AutoFit();

                shortHeaderWidth = spreadsheet?.document.GetColumnWidth("A1");
                longHeaderWidth = spreadsheet?.document.GetColumnWidth("B1");
            }

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename2))
            {
                spreadsheet.Write(AutoFit.Values);
                var shortColumnPrevWidth = spreadsheet?.document.GetColumnWidth("A1");
                var longColumnPrevWidth = spreadsheet?.document.GetColumnWidth("B1");

                spreadsheet?.AutoFit();

                var shortColumnWidth = spreadsheet?.document.GetColumnWidth("A1");
                var longColumnWidth = spreadsheet?.document.GetColumnWidth("B1");

                Assert.Multiple(() =>
                {
                    Assert.That(shortColumnWidth, Is.LessThan(shortColumnPrevWidth));
                    Assert.That(longColumnWidth, Is.GreaterThan(longColumnPrevWidth));
                });

                Assert.Multiple(() =>
                {
                    Assert.That(shortColumnWidth, Is.GreaterThan(shortHeaderWidth));
                    Assert.That(longColumnWidth, Is.GreaterThan(longHeaderWidth));
                });
            }
        }

        [Test]
        public void Write_WithOneSimpleObject_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three);

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet?.document.GetCellValueAsString("A1"), Is.EqualTo(nameof(Simple.Text)));
                Assert.That(spreadsheet?.document.GetCellValueAsString("A2"), Is.EqualTo("1"));
                Assert.That(spreadsheet?.document.GetCellValueAsString("A3"), Is.EqualTo("2"));
                Assert.That(spreadsheet?.document.GetCellValueAsString("A4"), Is.EqualTo("3"));
            });
        }

        [Test]
        public void Write_WithMultipleObjects_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three);
            spreadsheet.Write(AutoFit.Values);

            Assert.That(spreadsheet?.document.GetCurrentWorksheetName(), Is.EqualTo(nameof(AutoFit)));

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet.document.GetCellValueAsString("A1"), Is.EqualTo(nameof(AutoFit.S)));
                Assert.That(spreadsheet.document.GetCellValueAsString("B1"), Is.EqualTo(nameof(AutoFit.LongLongLongLong)));
                Assert.That(spreadsheet.document.GetCellValueAsString("A2"), Is.EqualTo("Short"));
            });

            spreadsheet.document.SelectWorksheet(nameof(Simple));

            Assert.That(spreadsheet.document.GetCurrentWorksheetName(), Is.EqualTo(nameof(Simple)));

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet.document.GetCellValueAsString("A1"), Is.EqualTo(nameof(Simple.Text)));
                Assert.That(spreadsheet.document.GetCellValueAsString("A2"), Is.EqualTo("1"));
                Assert.That(spreadsheet.document.GetCellValueAsString("A3"), Is.EqualTo("2"));
                Assert.That(spreadsheet.document.GetCellValueAsString("A4"), Is.EqualTo("3"));
            });

            Assert.That(spreadsheet.document.GetWorksheetNames(), Has.Count.EqualTo(2));
        }
    }
}