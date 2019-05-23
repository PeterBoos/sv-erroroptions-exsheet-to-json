using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace excel_sheet_to_json
{
    public class Options
    {
        [JsonProperty(PropertyName = "areas")]
        public List<Area> Areas { get; set; }
    }
}
