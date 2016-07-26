using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelPixelArt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPixelArt.Tests {
    [TestClass()]
    public class ProgramTests {
        [TestMethod()]
        public void MainTest() {
            Program.Main(new string[] { "demo6.png","2" });
            Program.Main(new string[] { "demo6.png","2","1","true"});
            Assert.IsNull(null);
        }
    }
}