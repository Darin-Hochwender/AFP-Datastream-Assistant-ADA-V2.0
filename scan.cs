using System;
using System.IO;

class Program {
    static void Main() {
        byte[] data = File.ReadAllBytes(@"C:\Users\Darin.Hochwender\_projects\AFP-Datastream-Assistant-ADA-V2.0\AFPEngineer_Explorer\AFP Files\DEV_IDS_150120261659_0092_Print.afp");
        for(int i=0; i<data.Length-8; i++) {
            if(data[i] == 0x5A) {
                int l = (data[i-2]<<8) | data[i-1];
                if(data[i+1] == 0xD3 && data[i+2] == 0xAB && data[i+3] == 0x8A) {
                    Console.WriteLine("MCF offset " + i + " len " + l);
                    for(int j=8; j<Math.Min(l,100); j++) Console.Write(data[i+j].ToString("X2") + " ");
                    Console.WriteLine();
                }
                if(data[i+1] == 0xD3 && data[i+2] == 0xAF && data[i+3] == 0xC3) {
                    Console.WriteLine("IOB offset " + i + " len " + l);
                    for(int j=8; j<Math.Min(l,100); j++) Console.Write(data[i+j].ToString("X2") + " ");
                    Console.WriteLine();
                }
                i += l-3;
            }
        }
    }
}
