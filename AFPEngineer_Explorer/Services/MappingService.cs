using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using AFPEngineer_Explorer.Models;

namespace AFPEngineer_Explorer.Services
{
    public class MappingService
    {
        private static Lazy<MappingService> _instance = new Lazy<MappingService>(() => new MappingService());
        public static MappingService Instance => _instance.Value;

        public AfpDefinitions Definitions { get; private set; }
        private Dictionary<string, TripletHeaderInfo> _tripletsList = new Dictionary<string, TripletHeaderInfo>();
        private Dictionary<string, ControlSequenceInfo> _ptocaControlSequences = new Dictionary<string, ControlSequenceInfo>();
        private string _tripletLayoutsDir;

        public MappingService() { }

        public void LoadDefinitions(string jsonFilePath)
        {
            if (File.Exists(jsonFilePath))
            {
                var options = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true };
                string json = File.ReadAllText(jsonFilePath);
                Definitions = JsonSerializer.Deserialize<AfpDefinitions>(json, options);
            }

            string dir = Path.GetDirectoryName(jsonFilePath);
            string tripletsPath = Path.Combine(dir, "triplets.json");
            string ptxPath = Path.Combine(dir, "ptx_control_sequences.json");
            _tripletLayoutsDir = Path.Combine(dir, "triplet_layouts");

            if (File.Exists(tripletsPath))
            {
                var options = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true, PropertyNameCaseInsensitive = true };
                var tList = JsonSerializer.Deserialize<List<TripletHeaderInfo>>(File.ReadAllText(tripletsPath), options);
                if (tList != null)
                {
                    foreach (var t in tList)
                    {
                        _tripletsList[t.Id.ToUpper()] = t;
                    }
                }
            }

            if (File.Exists(ptxPath))
            {
                var options = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true, PropertyNameCaseInsensitive = true };
                try {
                    using (JsonDocument doc = JsonDocument.Parse(File.ReadAllText(ptxPath)))
                    {
                        if (doc.RootElement.TryGetProperty("controlSequences", out JsonElement seqs))
                        {
                            var list = JsonSerializer.Deserialize<List<ControlSequenceInfo>>(seqs.GetRawText(), options);
                            if (list != null)
                            {
                                foreach (var item in list)
                                {
                                    _ptocaControlSequences[item.HexId.ToUpper()] = item;
                                }
                            }
                        }
                    }
                } catch { }
            }
        }

        public SfDefinition GetDefinition(string sfId)
        {
            if (Definitions == null || string.IsNullOrEmpty(sfId)) return null;
            return Definitions.SfDefinitions?.FirstOrDefault(x => x.SfId.Equals(sfId, StringComparison.OrdinalIgnoreCase));
        }

        public ParsedSfData ParseSfData(string sfId, byte[] payload, Encoding textEncoding, int? dynamicRgLength = null)
        {
            var result = new ParsedSfData();
            var def = GetDefinition(sfId);
            if (def == null || def.Layout == null) return result;

            bool isRepeatingGroup = false;
            if (def.Definition != null && def.Definition.TryGetValue("repeatingGroup", out JsonElement rpVal))
            {
                if (rpVal.ValueKind == JsonValueKind.True) isRepeatingGroup = true;
            }

            int baseOffset = 0;
            int groupIndex = 1;

            while (baseOffset < payload.Length)
            {
                int groupConsumed = 0;
                int groupEndIdx = payload.Length;

                if (dynamicRgLength.HasValue && dynamicRgLength.Value > 0)
                {
                    groupEndIdx = baseOffset + dynamicRgLength.Value;
                }

                foreach (var field in def.Layout)
                {
                    if (dynamicRgLength.HasValue && groupConsumed >= dynamicRgLength.Value)
                        break; // Stop parsing fields if we've reached the dynamic length of the group

                    if (field.Offset == null) continue;

                    int begin = baseOffset + field.Offset.Begin;
                    int length = field.Offset.Length;
                    
                    string parsedValue = "";
                    string headerDisplayName = field.DisplayName ?? field.Name;
                    if (isRepeatingGroup) headerDisplayName = $"{headerDisplayName} [{groupIndex}]";

                    if (field.Type == "TRIPLETS")
                    {
                        if (begin < groupEndIdx) {
                            int consumed = ParseTriplets(payload, begin, result, textEncoding, groupEndIdx);
                            groupConsumed = Math.Max(groupConsumed, (begin - baseOffset) + consumed);
                        }
                        result.LayoutRows.Add(new SfLayoutRow
                        {
                            Offset = begin.ToString(),
                            Length = "",
                            Type = field.Type,
                            DisplayName = headerDisplayName,
                            Description = field.Description,
                            Value = string.IsNullOrWhiteSpace(field.Triplets) ? "See Triplets List Below" : $"Defined: {field.Triplets}",
                            Required = field.Required.HasValue ? (field.Required.Value ? "T" : "F") : ""
                        });
                        continue; 
                    }

                    if (field.Type == "PTOCA")
                    {
                        if (begin < groupEndIdx) {
                            int consumed = ParsePtoca(payload, begin, result, textEncoding, groupEndIdx);
                            groupConsumed = Math.Max(groupConsumed, (begin - baseOffset) + consumed);
                        }
                        result.LayoutRows.Add(new SfLayoutRow
                        {
                            Offset = begin.ToString(),
                            Length = "",
                            Type = field.Type,
                            DisplayName = headerDisplayName,
                            Description = field.Description,
                            Value = "See PTOCA Control Sequences Below",
                            Required = field.Required.HasValue ? (field.Required.Value ? "T" : "F") : ""
                        });
                        continue; 
                    }

                    if (field.Type == "REPEATING_GROUP" && sfId == "FNN")
                    {
                        // FNN Map section
                        try
                        {
                            // Find the first offset to know where the map ends
                            int firstOffset = payload.Length;
                            int tempIdx = begin;
                            while (tempIdx + 12 <= Math.Min(firstOffset, payload.Length))
                            {
                                int tsidOff = (payload[tempIdx + 8] << 24) | (payload[tempIdx + 9] << 16) | (payload[tempIdx + 10] << 8) | payload[tempIdx + 11];
                                if (tsidOff > 0 && tsidOff < firstOffset) firstOffset = tsidOff;
                                tempIdx += 12;
                            }
                            
                            int count = (firstOffset - begin) / 12;
                            string val = "";
                            int mapIdx = begin;
                            for (int i = 0; i < count; i++)
                            {
                                string cgrid = textEncoding.GetString(payload, mapIdx, 8);
                                int tsidOff = (payload[mapIdx + 8] << 24) | (payload[mapIdx + 9] << 16) | (payload[mapIdx + 10] << 8) | payload[mapIdx + 11];
                                val += $"GCGID: {cgrid}, Offset: {tsidOff}\n";
                                mapIdx += 12;
                            }
                            result.LayoutRows.Add(new SfLayoutRow { Offset = begin.ToString(), Length = (firstOffset - begin).ToString(), Type = field.Type, DisplayName = headerDisplayName, Description = field.Description, Value = val });
                            groupConsumed = Math.Max(groupConsumed, firstOffset - begin);
                        }
                        catch { }
                        continue;
                    }

                    if (sfId == "CPI" && field.Name == "UnicodeVal")
                    {
                        if (dynamicRgLength == 0x0A || dynamicRgLength == 0x0B) continue; // Skip Unicode value
                        else if (dynamicRgLength == 0xFE) begin = baseOffset + 10;
                        else if (dynamicRgLength == 0xFF) begin = baseOffset + 11;
                    }

                    if (begin >= payload.Length)
                    {
                        if (field.Required.HasValue && !field.Required.Value)
                        {
                            parsedValue = "";
                        }
                        else
                        {
                            parsedValue = "[Offset Out of Bounds]";
                        }
                    }
                    else
                    {
                        if (sfId == "CPI" && field.Name == "CodePoint")
                        {
                            if (dynamicRgLength == 0x0A) length = 1;      // single byte
                            else if (dynamicRgLength == 0x0B) length = 2; // double byte
                            else if (dynamicRgLength == 0xFE) length = 1; // single byte + unicode
                            else if (dynamicRgLength == 0xFF) length = 2; // double byte + unicode
                        }

                        int lengthToRead = length;
                        if (lengthToRead == -1 || lengthToRead == 0)
                        {
                            lengthToRead = payload.Length - begin;
                            if (length == -1)
                            {
                                parsedValue = $"{lengthToRead} bytes of {field.DisplayName ?? field.Name}";
                            }
                        }

                        if (begin + lengthToRead > payload.Length) lengthToRead = payload.Length - begin;

                        if (string.IsNullOrEmpty(parsedValue))
                        {
                            byte[] fieldData = new byte[lengthToRead];
                            Array.Copy(payload, begin, fieldData, 0, lengthToRead);
                            parsedValue = DecodeFieldData(field.Type, fieldData, textEncoding);

                            if (field.EnumMap != null)
                            {
                                string hexKey = BitConverter.ToString(fieldData).Replace("-", "");
                                if (field.EnumMap.TryGetValue(hexKey, out string enumVal))
                                {
                                    parsedValue = $"{parsedValue} - {enumVal}";
                                }
                            }
                        }
                        groupConsumed = Math.Max(groupConsumed, (begin - baseOffset) + lengthToRead);
                        
                        if (isRepeatingGroup && (field.DisplayName == "RGLength" || field.Name == "RGLength" || field.Name == "PMCid") && int.TryParse(parsedValue, out int rgLen) && rgLen > 0)
                        {
                            groupEndIdx = Math.Min(payload.Length, baseOffset + rgLen);
                            groupConsumed = Math.Max(groupConsumed, rgLen);
                        }
                    }

                    result.LayoutRows.Add(new SfLayoutRow
                    {
                        Offset = begin.ToString(),
                        Length = length == -1 ? "-1" : length.ToString(),
                        Type = field.Type,
                        DisplayName = headerDisplayName,
                        Description = field.Description,
                        Value = parsedValue,
                        Required = field.Required.HasValue ? (field.Required.Value ? "T" : "F") : ""
                    });
                }

                if (!isRepeatingGroup) break;
                
                // If a dynamic length was provided by state (e.g., FNC for FNI), use it to advance.
                if (dynamicRgLength.HasValue && dynamicRgLength.Value > 0)
                {
                    groupConsumed = dynamicRgLength.Value;
                }

                if (groupConsumed <= 0) break;

                baseOffset += groupConsumed;
                groupIndex++;
            }

            return result;
        }

        private int ParsePtoca(byte[] payload, int startIdx, ParsedSfData result, Encoding textEncoding, int endIdx = -1)
        {
            if (endIdx == -1 || endIdx > payload.Length) endIdx = payload.Length;
            int n = endIdx;
            if (payload == null || n <= startIdx) return 0;
            
            int pos = startIdx;

            while (pos < n)
            {
                // Fallback check to skip 2B D3 if it's somehow there
                if (pos + 1 < n && payload[pos] == 0x2B && payload[pos + 1] == 0xD3)
                {
                    pos += 2;
                    continue;
                }

                byte lenByte = payload[pos];
                int len = lenByte == 0 ? 256 : (int)lenByte;
                if (len < 2) { pos++; continue; }
                
                int regionEnd = pos + len;
                if (regionEnd > n) break; // truncated

                byte csCode = payload[pos + 1];
                int dataStart = pos + 2;
                int dataLen = len - 2;

                string csId = csCode.ToString("X2");
                string csName = $"Unknown Control Sequence (0x{csId})";
                string url = null;
                
                if (_ptocaControlSequences.TryGetValue(csId, out var info))
                {
                    csName = $"{info.Name} [{info.ShortId}] (0x{csId})";
                    url = info.Url;
                }

                var parsedCs = new ParsedTriplet { HexId = csId, Heading = csName, Url = url };

                try
                {
                    byte[] csData = new byte[dataLen];
                    if (dataLen > 0)
                        Array.Copy(payload, dataStart, csData, 0, dataLen);

                    // Decode recognizable ones
                    string displayValue = "";
                    if (csCode == 0xD3 && dataLen >= 2)
                        displayValue = ((short)((csData[0] << 8) | csData[1])).ToString();
                    else if (csCode == 0xC7 && dataLen >= 2)
                        displayValue = ((short)((csData[0] << 8) | csData[1])).ToString();
                    else if (csCode == 0xDB)
                    {
                        if (dataLen > 0)
                        {
                            try { displayValue = textEncoding.GetString(csData).TrimEnd(); } 
                            catch { displayValue = "[Text Decode Error]"; }
                        }
                    }
                    else if (csCode == 0xF3 && dataLen >= 1)
                        displayValue = $"LID=0x{csData[0]:X2}";
                    else if (csCode == 0xE7 && dataLen >= 4)
                        displayValue = $"RLENGTH={(short)((csData[0]<<8)|csData[1])} RWIDTH={(short)((csData[2]<<8)|csData[3])}";
                    else if (csCode == 0xE5 && dataLen >= 4)
                        displayValue = $"RLENGTH={(short)((csData[0]<<8)|csData[1])} RWIDTH={(short)((csData[2]<<8)|csData[3])}";
                    else if (csCode == 0xF1 && dataLen >= 1)
                        displayValue = csData[0].ToString();
                    else if (csCode == 0x81 && dataLen >= 2)
                    {
                        byte cp = csData[1];
                        displayValue = cp switch { 0x01=>"RGB", 0x04=>"CMYK", 0x06=>"Highlight", 0x08=>"CIELAB", 0x40=>"Standard OCA", _=>$"Unknown(0x{cp:X2})" };
                        displayValue = $"Color space {displayValue}";
                    }
                    else if (csCode == 0xF7 && dataLen >= 4)
                    {
                        ushort v1 = (ushort)((csData[0] << 8) | csData[1]);
                        ushort v2 = (ushort)((csData[2] << 8) | csData[3]);
                        displayValue = $"Inline: {v1 & 0x01FF} deg, Baseline: {v2 & 0x01FF} deg";
                    }
                    else if (dataLen > 0)
                        displayValue = BitConverter.ToString(csData).Replace("-", "");

                    // Use the PtxParser's "check back soon" approach for the requested control sequences
                    if (csCode == 0xC1 || csCode == 0xC3 || csCode == 0xC9 || csCode == 0xD1 || csCode == 0xD5 || csCode == 0x6D || csCode == 0x8B || csCode == 0x8F || csCode == 0x99)
                        displayValue = "Work in progress - check back soon";

                    parsedCs.Rows.Add(new TripletRow { Offset = dataStart.ToString(), FieldDescription = "Value", Value = displayValue });
                    result.Triplets.Add(parsedCs);
                }
                catch { }

                pos = regionEnd;
            }

            return pos - startIdx;
        }

        private int ParseTriplets(byte[] payload, int startIdx, ParsedSfData result, Encoding textEncoding, int endIdx = -1)
        {
            if (endIdx == -1 || endIdx > payload.Length) endIdx = payload.Length;
            int idx = startIdx;
            while (idx < endIdx)
            {
                int len = payload[idx];
                if (len < 2 || idx + len > endIdx) break;

                string tId = payload[idx + 1].ToString("X2");
                byte[] tData = new byte[len];
                Array.Copy(payload, idx, tData, 0, len);

                var parsedT = new ParsedTriplet { HexId = tId };
                
                if (_tripletsList.TryGetValue(tId, out var info))
                {
                    parsedT.Heading = $"{info.Name} (0x{tId})";
                    parsedT.Url = info.Url;
                }
                else
                {
                    parsedT.Heading = $"Unknown Triplet (0x{tId})";
                }

                string jsonPath = Path.Combine(_tripletLayoutsDir, $"{tId}.json");
                if (File.Exists(jsonPath))
                {
                    string rawJson = File.ReadAllText(jsonPath);
                    parsedT.RawJson = rawJson;
                    
                    try
                    {
                        var options = new JsonSerializerOptions { ReadCommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true, PropertyNameCaseInsensitive = true };
                        var tDef = JsonSerializer.Deserialize<TripletDefinition>(rawJson, options);
                        if (tDef != null && tDef.Elements != null)
                        {
                            foreach (var el in tDef.Elements)
                            {
                                int elLen = 0;
                                if (el.Length.ValueKind == JsonValueKind.Number) elLen = el.Length.GetInt32();
                                else if (el.Length.ValueKind == JsonValueKind.String && el.Length.GetString() == "variable") 
                                    elLen = len - el.Offset;

                                int elBegin = el.Offset;
                                if (elBegin >= len) continue;

                                if (elBegin + elLen > len) elLen = len - elBegin;

                                byte[] elData = new byte[elLen];
                                Array.Copy(tData, elBegin, elData, 0, elLen);

                                string val = DecodeFieldData(el.Type, elData, textEncoding);

                                if (el.EnumMap != null)
                                {
                                    string hexKey = BitConverter.ToString(elData).Replace("-", "");
                                    if (el.EnumMap.TryGetValue(hexKey, out string enumVal))
                                    {
                                        val = $"{val} - {enumVal}";
                                    }
                                }

                                parsedT.Rows.Add(new TripletRow
                                {
                                    Offset = elBegin.ToString(),
                                    FieldDescription = el.Description ?? el.Name,
                                    Value = val
                                });
                            }
                        }
                    }
                    catch { }
                }

                if (parsedT.Rows.Count == 0)
                {
                    int rl = len > 2 ? len - 2 : 0;
                    byte[] rd = new byte[rl];
                    if (rl > 0) Array.Copy(tData, 2, rd, 0, rl);
                    parsedT.Rows.Add(new TripletRow { Offset = "2", FieldDescription = "Raw Data", Value = BitConverter.ToString(rd).Replace("-", "") });
                }

                result.Triplets.Add(parsedT);
                idx += len;
            }
            return idx - startIdx;
        }

        private string DecodeFieldData(string type, byte[] data, Encoding textEncoding)
        {
            if (data == null || data.Length == 0) return "[Empty]";

            switch (type?.ToUpper())
            {
                case "CHAR":
                    string txt = textEncoding.GetString(data);
                    return System.Text.RegularExpressions.Regex.Replace(txt, @"[^\u0020-\u007E\u00A0-\u00FF]", ".");
                case "UBIN":
                case "CODE":
                case "ENUM":
                    return BitConverter.ToString(data).Replace("-", "");
                case "SBIN":
                    if (data.Length == 1) return ((sbyte)data[0]).ToString();
                    if (data.Length == 2)
                    {
                        var temp = new byte[2];
                        Array.Copy(data, temp, 2);
                        if (BitConverter.IsLittleEndian) Array.Reverse(temp);
                        return BitConverter.ToInt16(temp, 0).ToString();
                    }
                    return BitConverter.ToString(data).Replace("-", "");
                case "BITS":
                    return Convert.ToString(data[0], 2).PadLeft(8, '0');
            }
            return BitConverter.ToString(data).Replace("-", " ");
        }
    }
}
