using System;
using System.Collections.Generic;
using System.Linq;

namespace ColumnMapper {
    public class ExcelColumnMapper {

        public ExcelColumnMapper() {

            // 컬럼에 표기될 심볼 테이블
            symbolTable = new [] { 
                "Not Zero Based Index",
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", 
                "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", 
                "U", "V", "W", "X", "Y", "Z",
            }.ToList();            
            
            // 사용될 심볼 개수
            symbolCount = symbolTable.Count - 1;

            // Build Short Range Cache
            shortRangeEnd = 702;
            shortRangeCache = new List<string>(shortRangeEnd + 1);
            shortRangeCache.Add("Not Zero Based Index");

            for(int i=1; i<shortRangeEnd+1; ++i) {
                var cn = ComputeColumnName(i);
                shortRangeCache.Add(cn);
            }
        }



        #region Member Field Group

        private readonly int symbolCount;
        private readonly int shortRangeEnd;
        private readonly List<string> symbolTable;
        private readonly List<string> shortRangeCache;

        #endregion



        #region Core Interface Method Group

        public string GetColumnName(int colnum) {

            if(colnum <= 0)
                throw new ArgumentOutOfRangeException();

            if(colnum <= shortRangeEnd)
                return shortRangeCache[colnum];     

            // IF Cache Hit Cache(colnum) or GetNext(Cache(colnum-1))

            return ComputeColumnName2(colnum);
        }

        #endregion



        #region Core Internal Method Group

        private string ComputeColumnName(int colnum) {

            if(colnum <= 0)  
                throw new ArgumentOutOfRangeException();

            if(colnum <= symbolCount)
                return symbolTable[colnum];

            int remain;
            int quotient = Math.DivRem(colnum, symbolCount, out remain);

            // 2 이상의 자리수 n이 있을 때 n 위치의 수치값 0은 
            // n-1 자리의 'Z'를 의미한다. 즉, A0는 Z로 표현된다. 
            if(remain == 0) {
                --quotient; 
                remain = symbolCount;
            }

            return ComputeColumnName(quotient) + ComputeColumnName(remain);
        }

        private string ComputeColumnName2(int colnum) {

            if(colnum == 0)
                return symbolTable[symbolCount];

            if(colnum <= symbolCount)
                return symbolTable[colnum];

            int remain;
            int quotient = Math.DivRem(colnum, symbolCount, out remain);

            // 2 이상의 자리수 n이 있을 때 n 위치의 수치값 0은 
            // n-1 자리의 'Z'를 의미한다. 즉, A0는 Z로 표현된다. 
            if(remain == 0) 
                --quotient;

            return ComputeColumnName2(quotient) + ComputeColumnName2(remain);
        }

        #endregion

    }
}
