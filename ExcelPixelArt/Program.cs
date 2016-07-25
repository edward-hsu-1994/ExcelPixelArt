using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPixelArt {
    public class Program {
        static void Main(string[] args) {
            FileStream file = new FileStream($"{Guid.NewGuid()}.xlsx", FileMode.Create);

            new PixelArtConverter().Convert("demo.png").Write(file);
            file.Close();
        }
    }
}
