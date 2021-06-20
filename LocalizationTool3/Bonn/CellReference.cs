using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSKT.Bonn
{
    public readonly struct CellReference
    {
        public readonly uint columnIndex;
        public readonly uint rowIndex;
        public readonly string value;

        public CellReference(string value)
        {
            var c = 0;
            var r = 0;
            foreach (var it in value)
            {
                if (it >= 'A' && it <= 'Z')
                {
                    c *= 26;
                    c += it - 'A';
                }
                else
                {
                    r *= 10;
                    r += it - '1';
                }
            }
            columnIndex = (uint)c + 1;
            rowIndex = (uint)r + 1;

            this.value = value;
        }

        public CellReference(uint column, uint row)
        {
            columnIndex = column;
            rowIndex = row;
            var columnName = "";
            --column;
            while (true)
            {
                var c = (char)('A' + (char)(column % 26));
                columnName = c.ToString() + columnName;

                column /= 26;
                if (column == 0)
                {
                    break;
                }
            }

            value = columnName + row.ToString();
        }
    }
}
