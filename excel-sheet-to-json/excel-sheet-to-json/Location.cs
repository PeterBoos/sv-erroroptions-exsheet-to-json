using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class Location
    {
        [JsonProperty(PropertyName = "caption")]
        public string Caption { get; set; }

        [JsonProperty(PropertyName = "Spaces")]
        public List<Space> Spaces { get; set; }
    }
}
