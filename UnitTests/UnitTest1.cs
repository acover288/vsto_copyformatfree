using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using vsto_copyformatfree;

namespace UnitTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void Test2NewLine()
        {
            string result = ThisAddIn.filterBody("a\r\n\r\nb");
            Assert.AreEqual("a\r\nb", result);
        }
        [TestMethod]
        public void Test3NewLine()
        {
            string result = ThisAddIn.filterBody("a\r\n\r\n\r\nb");
            Assert.AreEqual("a\r\nb", result);
        }
        [TestMethod]
        public void TestNoR()
        {
            string result = ThisAddIn.filterBody("a\r\n\n\r\nb");
            Assert.AreEqual("a\r\nb", result);
        }
    }
}
