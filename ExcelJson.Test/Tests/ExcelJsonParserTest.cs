using System.Text;
using System.Globalization;
using Newtonsoft.Json;
using FluentAssertions;
using FluentAssertions.Json;

namespace ExcelJson
{
    public class ExcelJsonParserTest
    {
        static readonly ExcelJsonParser m_Parser = new(new JsonSerializerSettings
        {
            Culture = CultureInfo.CurrentCulture,
        });

        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        static IEnumerable<ExcelJsonSheet> ReadExcel(string name)
        {
            var path = Path.Combine("../../../Excels/", name);
            using var stream = File.OpenRead(path);
            foreach (var sheet in m_Parser.ParseExcel(stream))
            {
                yield return sheet;
            }
        }

        [Test]
        public void FilterColumnTest()
        {
            var value = Enumerable.Range(1, 10)
                .Select(x => new ParsingTest1 { A = x })
                .ToDictionary(x => x.A);

            var actual = ReadExcel("FilterColumnTest.xlsx").First().Json;
            var expected = JsonConvert.SerializeObject(value);

            expected.Should().BeEquivalentTo(actual);
        }

        [Test]
        public void EmptyTest()
        {
            var value = Enumerable.Range(0, 0)
                .Select(x => new ParsingTest1 { A= x }) 
                .ToDictionary(x => x.A);

            var actual = ReadExcel("EmptyTest.xlsx").First().Json;
            var expected = JsonConvert.SerializeObject(value);

            actual.Should().BeEquivalentTo(expected);
        }

        [Test]
        public void ParsingTest1()
        {
            var value1 = Enumerable.Range(1, 10)
                .Select(x => new ParsingTest2 { A = x, B = x })
                .ToDictionary(x => x.A);

            var value2 = Enumerable.Range(1, 10)
                .Select(x => new ParsingTest1 { A = x })
                .ToDictionary(x => x.A);

            var actual = ReadExcel("ParsingTest1.xlsx").First().Json;
            var expected1 = JsonConvert.SerializeObject(value1);
            var expected2 = JsonConvert.SerializeObject(value2);

            expected1.Should().NotBeEquivalentTo(actual);
            expected2.Should().BeEquivalentTo(actual);
        }

        [Test]
        public void ParsingTest2()
        {
            var value = Enumerable.Range(1, 10)
                .Select(x => new ParsingTest2 { A = x, B = x })
                .ToDictionary(x => x.A);

            var actual = ReadExcel("ParsingTest2.xlsx").First().Json;
            var expected = JsonConvert.SerializeObject(value);

            expected.Should().BeEquivalentTo(actual);
        }

        [Test]
        public void ParsingTest3()
        {
            var json = ReadExcel("ParsingTest3.xlsx").First().Json;
            var actual = JsonConvert.DeserializeObject<Dictionary<int, ParsingTest3>>(json);
            actual.Should().NotBeNullOrEmpty();
        }

        [Test]
        public void PrimitiveTypeTest()
        {
            var actual = ReadExcel("PrimitiveTypeTest.xlsx").First().Json;
            actual.Should().NotBeNullOrEmpty();
        }

        [Test]
        public void FilterSheetTest()
        {
            var actual = ReadExcel("FilterSheetTest.xlsx");
            actual.Should().HaveCount(0);
        }

        [Test]
        public void ArrayTest()
        {
            var value = Enumerable.Range(1, 5)
                .Select(x => new ArrayTestSheet1 { A = [x * 10, x, x] })
                .ToDictionary(x => x.A == null ? -1 : x.A[0]);

            var actual = ReadExcel("ArrayTest.xlsx").First().Json;
            var expected = JsonConvert.SerializeObject(value);

            expected.Should().BeEquivalentTo(actual);
        }

        [Test]
        public void ArrayPlaceholderExceptionTest()
        {
            try
            {
                var actual = ReadExcel("ArrayPlaceholderExceptionTest.xlsx").First().Json;
                Assert.Fail(actual);
            }
            catch (Exception ex) 
            {
                (ex is ArrayPlaceholderException).Should()
                    .BeTrue();
            }
        }

        [Test]
        public void DuplicatedExceptionTest()
        {
            try
            {
                var actual = ReadExcel("DuplicatedExceptionTest.xlsx").First().Json;
                Assert.Fail(actual);
            }
            catch (Exception ex)
            {
                (ex is DuplicatedException).Should()
                    .BeTrue();
            }
        }

        [Test]
        public void TooManyTypeTokenExceptionTest()
        {
            try
            {
                var actual = ReadExcel("TooManyTypeTokenExceptionTest.xlsx").First().Json;
                Assert.Fail(actual);
            }
            catch (Exception ex)
            {
                if (ex is TypeTokenException typeTokenException)
                {
                    typeTokenException.Definition
                        .Count(x => x == typeTokenException.TypeToken)
                        .Should()
                        .NotBe(1);
                }
            }
        }

        [Test]
        public void MissingTypeTokenExceptionTest()
        {
            try
            {
                var actual = ReadExcel("MissingTypeTokenExceptionTest.xlsx").First().Json;
                Assert.Fail(actual);
            }
            catch (Exception ex)
            {
                if (ex is TypeTokenException typeTokenException)
                {
                    typeTokenException.Definition
                        .Count(x => x == typeTokenException.TypeToken)
                        .Should()
                        .Be(0);
                }
            }
        }
    }

    public class ParsingTest1
    {
        public int A { get; set; }
    }

    public class ParsingTest2
    {
        public int A { get; set; }
        public int B { get; set; }
    }

    public class ParsingTest3
    {
        public int Id { get; set; }
        public DateTime OrderDate { get; set; }
        public string? Region { get; set; }
        public string? Rep { get; set; }
        public string? Item { get; set; }
        public int Units { get; set; }
        public float UnitCost { get; set; }
        public float Total { get; set; }
    }

    public class ArrayTestSheet1
    {
        public int[]? A { get; set; }
    }
}