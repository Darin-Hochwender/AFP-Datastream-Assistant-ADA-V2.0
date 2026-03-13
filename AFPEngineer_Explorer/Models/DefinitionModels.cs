using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AFPEngineer_Explorer.Models
{
    public class AfpDefinitions
    {
        [JsonPropertyName("schemaVersion")]
        public string SchemaVersion { get; set; }

        [JsonPropertyName("lastUpdated")]
        public string LastUpdated { get; set; }

        [JsonPropertyName("sfDefinitions")]
        public List<SfDefinition> SfDefinitions { get; set; }
    }

    public class SfDefinition
    {
        [JsonPropertyName("sfId")]
        public string SfId { get; set; }

        [JsonPropertyName("sfName")]
        public string SfName { get; set; }

        [JsonPropertyName("category")]
        public string Category { get; set; }

        [JsonPropertyName("definition")]
        public Dictionary<string, JsonElement> Definition { get; set; }

        [JsonPropertyName("layout")]
        public List<AfpLayoutField> Layout { get; set; }

        [JsonPropertyName("comments")]
        public string Comments { get; set; }
    }

    public class AfpLayoutField
    {
        [JsonPropertyName("offset")]
        public AfpOffset Offset { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("displayName")]
        public string DisplayName { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("triplets")]
        public string Triplets { get; set; }

        [JsonPropertyName("enumMap")]
        public Dictionary<string, string> EnumMap { get; set; }

        [JsonPropertyName("required")]
        public bool? Required { get; set; }
    }

    public class AfpOffset
    {
        [JsonPropertyName("begin")]
        public int Begin { get; set; }

        [JsonPropertyName("length")]
        public int Length { get; set; }
    }

    public class TripletHeaderInfo
    {
        [JsonPropertyName("id")]
        public string Id { get; set; }
        
        [JsonPropertyName("name")]
        public string Name { get; set; }
        
        [JsonPropertyName("url")]
        public string Url { get; set; }
    }

    public class TripletDefinition
    {
        [JsonPropertyName("elements")]
        public List<TripletElement> Elements { get; set; }
    }

    public class TripletElement
    {
        [JsonPropertyName("offset")]
        public int Offset { get; set; }

        [JsonPropertyName("length")]
        public JsonElement Length { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("enumMap")]
        public Dictionary<string, string> EnumMap { get; set; }
    }

    // Parsed models for UI binding
    public class ParsedSfData
    {
        public List<SfLayoutRow> LayoutRows { get; set; } = new List<SfLayoutRow>();
        public List<ParsedTriplet> Triplets { get; set; } = new List<ParsedTriplet>();
    }

    public class SfLayoutRow
    {
        public string Offset { get; set; }
        public string Length { get; set; }
        public string Type { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string Value { get; set; }
        public string Required { get; set; }
    }

    public class ParsedTriplet
    {
        public string HexId { get; set; }
        public string Heading { get; set; }
        public string Url { get; set; }
        public string RawJson { get; set; }
        public List<TripletRow> Rows { get; set; } = new List<TripletRow>();
    }

    public class TripletRow
    {
        public string Offset { get; set; }
        public string FieldDescription { get; set; }
        public string Value { get; set; }
    }
    public class ControlSequenceInfo
    {
        [JsonPropertyName("hexId")]
        public string HexId { get; set; }

        [JsonPropertyName("shortId")]
        public string ShortId { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("url")]
        public string Url { get; set; }
    }}
