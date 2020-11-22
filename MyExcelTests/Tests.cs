using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyExcel;
using System;
using System.Collections.Generic;
using System.Text;

namespace MyExcel.Tests
{
    [TestClass()]
    public class Tests
    {
        public static Cell temp = new Cell();
        public static Dictionary<string, Cell> dict = new Dictionary<string, Cell>() { { "A1", temp } };
        public Calculator calculator = new Calculator(dict);
        string dataLost = "";
        [TestMethod()]
        public void EvaluateTest()          // тестування зчитування виразу
        {
            string expr = "mmin(mmax(5,10,3,15),10,25)";
            double expectedRes = 10;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void RecursiveTest()        // тестування рекурсії
        {
            string expr = "A1";  //вираз в А1
            double res1 = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(dict["A1"].Val, "ErrorCycle");
        }

        [TestMethod()]
        public void DivideZeroTest()      // тестування ділення на нуль
        {
            string expr = "5/0";  //вираз
            double res1 = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(dict["A1"].Val, "ErrorDivZero");
        }
    }
}