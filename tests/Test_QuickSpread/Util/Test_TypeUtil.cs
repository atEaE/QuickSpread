using QuickSpread.Util;
using System;
using Xunit;

namespace Test_QuickSpread.Util
{
    public class Test_TypeUtil
    {
        #region String
        [Fact]
        public void TestIsPrimitiveMethod_String()
        {
            string value = "Sample";
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }
        #endregion

        #region int
        [Fact]
        public void TestIsPrimitiveMethod_Int16()
        {
            Int16 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }

        [Fact]
        public void TestIsPrimitiveMethod_Int32()
        {
            Int32 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }

        [Fact]
        public void TestIsPrimitiveMethod_Int64()
        {
            Int64 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }
        #endregion

        #region uint
        [Fact]
        public void TestIsPrimitiveMethod_UInt16()
        {
            UInt16 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }

        [Fact]
        public void TestIsPrimitiveMethod_UInt32()
        {
            UInt32 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }

        [Fact]
        public void TestIsPrimitiveMethod_UInt64()
        {
            UInt64 value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }
        #endregion

        #region double
        [Fact]
        public void TestIsPrimitiveMethod_double()
        {
            double value = 1.0;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }
        #endregion

        #region float
        [Fact]
        public void TestIsPrimitiveMethod_float()
        {
            float value = 1;
            TypeUtil.IsPrimitive(value.GetType()).IsTrue();
        }
        #endregion
    }
}
