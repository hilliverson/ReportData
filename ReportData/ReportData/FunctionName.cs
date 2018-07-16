using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportData
{
    class FunctionName
    {
        private List<int> TheList;
        private readonly int FLAGGED = 4;

        public List<int> GetThem()
        {
            List<int> list1 = new List<int>();
            foreach (var x in TheList)
            {
                if (x == 4)
                {
                    list1.Add(x);
                }
            }

            return list1;
        }

        public List<int> GetFlaggedCells()
        {
            List<int> FlaggedCells = new List<int>();
            foreach (var cell in TheList)
            {
                if (cell == FLAGGED)
                {
                    FlaggedCells.Add(cell);
                }
            }
            return FlaggedCells;
        }
    }
}
