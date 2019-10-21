using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Demo_UT
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Demo_EPPlus.Test test = new Demo_EPPlus.Test();
            test.createFile(@"test.xlsx");
        }
    }
}
