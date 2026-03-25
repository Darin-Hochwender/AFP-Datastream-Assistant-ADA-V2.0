using System;
using System.IO;

class Program {
    static void Main() {
        byte[] data = File.ReadAllBytes(@"C:\Users\Darin.Hochwender\_projects\AFP-Datastream-Assistant-ADA-V2.0\AFPEngineer_Explorer\AFP Files\DEV_IDS_150120261659_0092_Print.afp");
        using (var w = new StreamWriter("output_dump.txt")) {
            for(int i=0; i<data.Length-8; i++) {
                if(data[i] == 0x5A) {
                    int l = (data[i-2]<<8) | data[i-1];
                    if(data[i+1] == 0xD3 && data[i+2] == 0xAB && data[i+3] == 0x8A) {
                        w.Write("MCF len " + l + ": ");
                        for(int j=8; j<Math.Min(l,200); j++) w.Write(data[i+j].ToString("X2") + " ");
                        w.WriteLine();
                    }
                    if(data[i+1] == 0xD3 && data[i+2] == 0xAF && data[i+3] == 0xC3) {
                        w.Write("IOB len " + l + ": ");
                        for(int j=8; j<Math.Min(l,300); j++) w.Write(data[i+j].ToString("X2") + " ");
                        w.WriteLine();
                    }
                    i += l-3;
                }
            }
        }
    }
}