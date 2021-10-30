using Microsoft.VisualStudio.TestTools.UnitTesting;
using Laba1Excel;

namespace UnitTestParser
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string x = "(15/3+2)*2";
            double expected = 14;

            Parser1 parser = new Parser1();
            double actual = parser.Evaluate(x);

            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod2()
        {
            string x = "(18/3+2)*7/2 div 5";
            double expected = 5;

            Parser1 parser = new Parser1();
            double actual = parser.Evaluate(x);

            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod3()
        {
            string x = "(2^5 div 5) mod 2^2";
            double expected = 2;

            Parser1 parser = new Parser1();
            double actual = parser.Evaluate(x);

            Assert.AreEqual(expected, actual);
        }
    }
}
