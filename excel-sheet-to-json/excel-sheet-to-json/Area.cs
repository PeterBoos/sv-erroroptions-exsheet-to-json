using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class Area

    {
        [JsonProperty(PropertyName = "caption")]
        public string Caption { get; set; }

        [JsonProperty(PropertyName = "code")]
        public string Code { get; set; }

        [JsonProperty(PropertyName = "locations")]
        public List<Location> Locations { get; set; }
    }
}
