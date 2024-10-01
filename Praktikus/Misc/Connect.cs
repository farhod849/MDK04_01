using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Praktikus.Misc
{
    internal class Connect
    {
        public static PraktikusBDEntities c;
        public static PraktikusBDEntities context
        {
            get
            {
                if (c == null)
                {
                    c = new PraktikusBDEntities();
                }
                return c;
            }
        }
    }
}
