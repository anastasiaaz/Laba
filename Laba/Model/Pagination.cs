using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laba.Model
{
    public class Pagination
    {
        public List<ShortNote> Data { get; }

        public int PageIndex { get; set; }
        public const int numberOfRecPerPage = 15;

        public Pagination()
        {
            Data = new List<ShortNote>();
            PageIndex = 1;
        }
    }
}
