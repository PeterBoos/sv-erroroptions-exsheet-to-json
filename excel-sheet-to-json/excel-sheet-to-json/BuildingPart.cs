using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class BuildingPart
    {
        [JsonProperty(PropertyName = "caption")]
        public string Caption { get; set; }

        [JsonProperty(PropertyName = "code")]
        public string Code { get; set; }

        //[JsonProperty(PropertyName = "work order")]
        //public WorkOrder WorkOrder { get; set; }
    }
}