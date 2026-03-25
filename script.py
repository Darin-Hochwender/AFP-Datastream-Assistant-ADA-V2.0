import os

data = open(r"C:\Users\Darin.Hochwender\_projects\AFP-Datastream-Assistant-ADA-V2.0\AFPEngineer_Explorer\AFP Files\DEV_IDS_150120261659_0092_Print.afp", "rb").read()

with open("output_dump.txt", "w") as w:
    i = 0
    while i < len(data) - 8:
        if data[i] == 0x5A:
            l = (data[i-2] << 8) | data[i-1]
            if data[i+1] == 0xD3 and data[i+2] == 0xAB and data[i+3] == 0x8A:
                w.write(f"MCF len {l}: ")
                hex_data = " ".join([f"{x:02X}" for x in data[i+8 : i+min(l, 200)]])
                w.write(hex_data + "\n")
            if data[i+1] == 0xD3 and data[i+2] == 0xAF and data[i+3] == 0xC3:
                w.write(f"IOB len {l}: ")
                hex_data = " ".join([f"{x:02X}" for x in data[i+8 : i+min(l, 300)]])
                w.write(hex_data + "\n")
            i += l - 3
        else:
            i += 1
