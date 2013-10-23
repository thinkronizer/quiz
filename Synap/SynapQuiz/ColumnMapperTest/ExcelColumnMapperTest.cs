using System;
using System.Collections.Generic;
using System.Linq;
using ColumnMapper;
using NUnit.Framework;

namespace ColumnMapperTest {
    [TestFixture]
    class ExcelColumnMapperTest {

        private ExcelColumnMapper mapper;
        private Dictionary<int,string> expects;

        [SetUp]
        public void SetUp() {
            mapper = new ExcelColumnMapper();
            expects = new Dictionary<int, string>();
        }



        #region Real Test Method Group

        [Test] 
        public void ArgumentBoundaryTest() {

            // Arrange : Notice! No zero-based index
            var invalidColNums = new int[] {0, -1, int.MinValue };

            // Act & Assert 
            foreach(int cn in invalidColNums) 
                Assert.Throws<ArgumentOutOfRangeException>(() => mapper.GetColumnName(cn));
        }

        [Test]
        public void InRangeTest() {

            // Arrange 
            // Append Test Set [A ~ AC]
            AppendTestSample(001, new [] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", 
                                           "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", 
                                           "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD" });
            // Append Test Set [AA ~ BA]
            AppendTestSample(027, new [] { "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", 
                                           "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", 
                                           "AU", "AV", "AW", "AX", "AY", "AZ", "BA" });
            // Append Test Set [ZX ~ AAB]
            AppendTestSample(700, new [] { "ZX", "ZY", "ZZ", "AAA", "AAB" });

            // Append Test Set [ZX ~ AAB]
            AppendTestSample(457677, new [] { "YZZY", "YZZZ", "ZAAA", "ZAAB", "ZAAC",});

            // Sequential Act & Assert
            foreach(var e in expects) {
                var actual = mapper.GetColumnName(e.Key);
                StringAssert.AreEqualIgnoringCase(e.Value, actual);
            }
        }

        #endregion



        #region Auxiliary Method Group

        private void AppendTestSample(int begin, string[] data) {

            var colnums = Enumerable.Range(begin, data.Length);

            for(int i=0; i < data.Length; ++i) {
                var colnum = colnums.ElementAt(i);
                expects[colnum] = data[i];
            }
        }

        #endregion

    }
}