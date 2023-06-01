using SpreadsheetUtility;

namespace Tests
{
    [TestFixture]
    public class SpreadsheetTests
    {
        const string k_SpreadsheetFilename = "test.xlsx";
        const string k_SpreadsheetFilename2 = "test2.xlsx";
        const string k_FolderName = "Folder/";
        const string k_NamedSheet = "NamedSheet";

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

        [Layout(Flow.Horizontal)]
        class Horizontal { }

        [Layout(Flow.Vertical)]
        class Vertical { }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(k_FolderName))
                Directory.Delete(k_FolderName, true);

            File.Delete(k_SpreadsheetFilename);
            File.Delete(k_SpreadsheetFilename2);
        }

        [TestCase(0, 0, typeof(Horizontal), "A1", TestName = "Cell(Horizontal): 0x0 -> A1")]
        [TestCase(6, 1, typeof(Horizontal), "G2", TestName = "Cell(Horizontal): 6x1 -> G2")]
        [TestCase(25, 0, typeof(Horizontal), "Z1", TestName = "Cell(Horizontal): 25x0 -> Z1")]
        [TestCase(26, 0, typeof(Horizontal), "AA1", TestName = "Cell(Horizontal): 26x0 -> AA1")]
        [TestCase(27, 0, typeof(Horizontal), "AB1", TestName = "Cell(Horizontal): 27x0 -> AB1")]
        [TestCase(30, 0, typeof(Horizontal), "AE1", TestName = "Cell(Horizontal): 30x0 -> AE1")]
        [TestCase(51, 0, typeof(Horizontal), "AZ1", TestName = "Cell(Horizontal): 51x0 -> AZ1")]
        [TestCase(52, 0, typeof(Horizontal), "BA1", TestName = "Cell(Horizontal): 52x0 -> BA1")]
        [TestCase(701, 0, typeof(Horizontal), "ZZ1", TestName = "Cell(Horizontal): 701x0 -> ZZ1")]
        [TestCase(702, 0, typeof(Horizontal), "AAA1", TestName = "Cell(Horizontal): 702x0 -> AAA1")]
        [TestCase(950, 950, typeof(Horizontal), "AJO951", TestName = "Cell(Horizontal): 950x950 -> AJO951")]
        [TestCase(2023, 2023, typeof(Horizontal), "BYV2024", TestName = "Cell(Horizontal): 2023x2023 -> BYV2024")]

        [TestCase(0, 1, typeof(Vertical), "B1", TestName = "Cell(Vertical): 0x1 -> B1")]
        [TestCase(6, 2, typeof(Vertical), "C7", TestName = "Cell(Vertical): 6x2 -> C7")]
        [TestCase(9, 25, typeof(Vertical), "Z10", TestName = "Cell(Vertical): 9x25 -> Z10")]
        [TestCase(10, 26, typeof(Vertical), "AA11", TestName = "Cell(Vertical): 10x26 -> AA11")]
        public void Cell_ConvertsXYToCellCoordinates(int x, int y, Type type, string expectedResult)
        {
            using var layout = new LayoutScope(type);
            Assert.That(Spreadsheet.Cell(x, y), Is.EqualTo(expectedResult));
        }

        [Test]
        public void Constructor_OpensOrCreatesSpreadsheet()
        {
            Assert.That(File.Exists(k_SpreadsheetFilename), Is.False);

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A1"), Is.Empty);

                spreadsheet.Write(Simple.One);

                Assert.That(spreadsheet?.Document.GetCellValueAsString("A1"), Is.Not.Empty);
            }

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A1"), Is.Not.Empty,
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

            var shortColumnPrevWidth = spreadsheet?.Document.GetColumnWidth("A1");
            var longColumnPrevWidth = spreadsheet?.Document.GetColumnWidth("B1");

            spreadsheet?.AutoFit();

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet?.Document.GetColumnWidth("A1"), Is.LessThan(shortColumnPrevWidth));
                Assert.That(spreadsheet?.Document.GetColumnWidth("B1"), Is.GreaterThan(longColumnPrevWidth));
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

                shortHeaderWidth = spreadsheet?.Document.GetColumnWidth("A1");
                longHeaderWidth = spreadsheet?.Document.GetColumnWidth("B1");
            }

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename2))
            {
                spreadsheet.Write(AutoFit.Values);
                var shortColumnPrevWidth = spreadsheet?.Document.GetColumnWidth("A1");
                var longColumnPrevWidth = spreadsheet?.Document.GetColumnWidth("B1");

                spreadsheet?.AutoFit();

                var shortColumnWidth = spreadsheet?.Document.GetColumnWidth("A1");
                var longColumnWidth = spreadsheet?.Document.GetColumnWidth("B1");

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
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A1"), Is.EqualTo(nameof(Simple.Text)));
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A2"), Is.EqualTo("1"));
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A3"), Is.EqualTo("2"));
                Assert.That(spreadsheet?.Document.GetCellValueAsString("A4"), Is.EqualTo("3"));
            });
        }

        [Test]
        public void Write_WithMultipleObjects_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three);
            spreadsheet.Write(AutoFit.Values);

            Assert.That(spreadsheet?.Document.GetCurrentWorksheetName(), Is.EqualTo(nameof(AutoFit)));

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet.Document.GetCellValueAsString("A1"), Is.EqualTo(nameof(AutoFit.S)));
                Assert.That(spreadsheet.Document.GetCellValueAsString("B1"), Is.EqualTo(nameof(AutoFit.LongLongLongLong)));
                Assert.That(spreadsheet.Document.GetCellValueAsString("A2"), Is.EqualTo("Short"));
            });

            spreadsheet.Document.SelectWorksheet(nameof(Simple));

            Assert.That(spreadsheet.Document.GetCurrentWorksheetName(), Is.EqualTo(nameof(Simple)));

            Assert.Multiple(() =>
            {
                Assert.That(spreadsheet.Document.GetCellValueAsString("A1"), Is.EqualTo(nameof(Simple.Text)));
                Assert.That(spreadsheet.Document.GetCellValueAsString("A2"), Is.EqualTo("1"));
                Assert.That(spreadsheet.Document.GetCellValueAsString("A3"), Is.EqualTo("2"));
                Assert.That(spreadsheet.Document.GetCellValueAsString("A4"), Is.EqualTo("3"));
            });

            Assert.That(spreadsheet.Document.GetWorksheetNames(), Has.Count.EqualTo(2));
        }

        [Test]
        public void Write_NamedSheet_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three, k_NamedSheet);

            Assert.That(spreadsheet.Document.GetWorksheetNames(), Contains.Item(k_NamedSheet));
        }

        [Test]
        public void Read_WithOneSimpleObject_Works()
        {
            Write_WithMultipleObjects_Works();

            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            var data = spreadsheet.Read<Simple>();

            Assert.Multiple(() =>
            {
                Assert.That(data, Is.Not.Null);
                Assert.That(data?.ElementAt(0).Text, Is.EqualTo("1"));
                Assert.That(data?.ElementAt(1).Text, Is.EqualTo("2"));
                Assert.That(data?.ElementAt(2).Text, Is.EqualTo("3"));
            });
        }

        [Test]
        public void Read_WithMultipleObjects_Works()
        {
            Write_WithMultipleObjects_Works();

            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            var simpleData = spreadsheet.Read<Simple>() as List<Simple>;
            var autoFitData = spreadsheet.Read<AutoFit>() as List<AutoFit>;

            Assert.Multiple(() =>
            {
                Assert.That(simpleData, Is.Not.Null);
                Assert.That(simpleData, Has.Count.EqualTo(Simple.Three.Length));
                Assert.That(simpleData?.ElementAt(0).Text, Is.EqualTo(Simple.Three[0].Text));
                Assert.That(simpleData?.ElementAt(1).Text, Is.EqualTo(Simple.Three[1].Text));
                Assert.That(simpleData?.ElementAt(2).Text, Is.EqualTo(Simple.Three[2].Text));
            });

            Assert.Multiple(() =>
            {
                Assert.That(autoFitData, Is.Not.Null);
                Assert.That(autoFitData, Has.Count.EqualTo(AutoFit.Values.Length));
                Assert.That(autoFitData?.ElementAt(0).S, Is.EqualTo(AutoFit.Values[0].S));
                Assert.That(autoFitData?.ElementAt(0).LongLongLongLong, Is.EqualTo(AutoFit.Values[0].LongLongLongLong));
            });
        }

        [Test]
        public void Read_NamedSheet_Works()
        {
            Write_NamedSheet_Works();

            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            var data = spreadsheet.Read<Simple>(k_NamedSheet);

            Assert.That(data, Is.Not.Null);
        }

        [Test]
        public void SetStartupSheet_IdentifiedByType_Works()
        {
            Write_WithMultipleObjects_Works();

            var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);

            Assert.That(spreadsheet.Document.GetCurrentWorksheetName(), Is.EqualTo(nameof(AutoFit)));

            spreadsheet.SetStartupSheet<Simple>();

            Assert.That(spreadsheet.Document.GetCurrentWorksheetName(), Is.EqualTo(nameof(AutoFit)));

            spreadsheet.Dispose();

            using var reopenedSpreadsheet = new Spreadsheet(k_SpreadsheetFilename);

            Assert.That(reopenedSpreadsheet.Document.GetCurrentWorksheetName(), Is.EqualTo(nameof(Simple)));
        }

        [Test]
        public void SetStartupSheet_IdentifiedByName_Works()
        {
            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                spreadsheet.Write(Simple.Three, k_NamedSheet);
                spreadsheet.Write(AutoFit.Values);
                spreadsheet.SetStartupSheet(k_NamedSheet);
            }

            using (var spreadsheet = new Spreadsheet(k_SpreadsheetFilename))
            {
                Assert.That(spreadsheet.Document.GetCurrentWorksheetName(), Is.EqualTo(k_NamedSheet));
            }
        }

        [Test]
        public void Delete_IdentifiedByType_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three);

            Assert.That(spreadsheet.Read<Simple>(), Is.Not.Null);

            spreadsheet.Delete<Simple>();

            Assert.That(spreadsheet.Read<Simple>(), Is.Null);
        }

        [Test]
        public void Delete_IdentifiedByName_Works()
        {
            using var spreadsheet = new Spreadsheet(k_SpreadsheetFilename);
            spreadsheet.Write(Simple.Three, k_NamedSheet);

            Assert.That(spreadsheet.Read<Simple>(k_NamedSheet), Is.Not.Null);

            spreadsheet.Delete(k_NamedSheet);

            Assert.That(spreadsheet.Read<Simple>(k_NamedSheet), Is.Null);
        }
    }
}