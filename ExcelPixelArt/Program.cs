using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPixelArt {
    public class Program {
        public static void Main(string[] args) {
            if(args?.Length == 0) {
                Console.WriteLine("請將圖片檔拖曳至這個EXE檔案放開，即會在圖片目錄產生EXCEL檔案");
                Console.WriteLine("命令列呼叫方式之參數為檔案路徑(必要)、EXCEL欄位大小、縮放層級(預設為1，2表示長寬除2)、色彩低品質(True/False)");
                Console.WriteLine("請按任意鍵，關閉本視窗...");
                Console.ReadKey();
            }
            
            short pSize = 4, cl = 1;
            bool lowQ = false;
            if(args.Length > 1) short.TryParse(args[1], out pSize);
            if(args.Length > 2) short.TryParse(args[2], out cl);
            if (args.Length > 3) bool.TryParse(args[3], out lowQ);
            new PixelArtConverter(pSize,cl,lowQ).Convert(args[0], $"{Guid.NewGuid()}.xlsx");
        }
    }
}
