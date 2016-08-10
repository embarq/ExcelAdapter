using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ExcelAdapter
{
    class JsonImport
    {
        public static List<List<KeyValuePair<string, string>>> From(string filename)
        {
            var data = new List<List<KeyValuePair<string, string>>>();
            dynamic jsonData = JArray.Parse(File.ReadAllText(filename));
            foreach (dynamic entry in jsonData)
            {
                var subSet = new List<KeyValuePair<string, string>>();
                foreach (dynamic property in entry)
                {
                    subSet.Add(new KeyValuePair<string, string>(property.Name, property.Value.ToString()));
                }
                data.Add(subSet);
            }
            return data;
        }
    }
}
