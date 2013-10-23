using System;
using ColumnMapper;

namespace ColumnMapperTest {
    class Program {
        public static void Main() {

            var mapper = new ExcelColumnMapper();
            Console.WriteLine(mapper.GetColumnName(99));
        }
    }
}
