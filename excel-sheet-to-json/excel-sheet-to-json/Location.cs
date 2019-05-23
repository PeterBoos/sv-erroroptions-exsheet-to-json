using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class Location
    {
        [JsonProperty(PropertyName = "caption")]
        public string Caption { get; set; }

        [JsonProperty(PropertyName = "code")]
        public string Code { get; set; }

        [JsonProperty(PropertyName = "building parts")]
        public List<BuildingPart> BuildingParts { get; set; }
    }
}