using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKeys
{
    public class Product
    {
        public string Name { get; set; }
        public string Key { get; set; }
        public string Id { get; set; }
        public bool IsNew { get; set; }
        public string Link { get; set; }
        public string Language { get; set; }
        public string Added { get; set; }
        public DateTime AddedDate { get; set; }
    }
}
