using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class WorkOrder
    {
        [JsonProperty(PropertyName = "type")]
        public string Type { get; set; }

        [JsonProperty(PropertyName = "category")]
        public string Category { get; set; }
    }
}