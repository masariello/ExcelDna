using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDna.IntegrationTests
{
    public class MarshalingTests
    {
        [ExcelFunction]
        public static double testAdd(double d1, double d2)
        {
            return d1+d2;
        }

        [ExcelFunction]
        public static double testMult(double d1, double d2)
        {
            return d1*d2;
        }

        [ExcelFunction]
        public static string testConcat(string s1, string s2)
        {
            return s1 + s2;
        }

        public enum TestEnum { negative, zero, positive };

        [ExcelFunction]
        public static double testEnumParam(TestEnum e)
        {
            double r = 0;
            switch (e)
            {
                case TestEnum.negative: { r = -1; break; }
                case TestEnum.zero: { r = 0; break; }
                case TestEnum.positive: { r = 1; break; }
            }
           return r;
        }

        [ExcelFunction]
        public static TestEnum testEnumResult(double d)
        {
            if (d < 0)
                return TestEnum.negative;
            else if (d > 0)
                return TestEnum.positive;
            else return TestEnum.zero;
        }
    }
}
