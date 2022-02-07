using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ManagedCom
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("240A2CF6-AABE-47E7-A163-3FF54FAFFFD0")]
    [ProgId("ManagedCom.SomeCom123")]
    public class nrComType
    {
        private static int iCount = 887;

        [ComVisible(true)]
        public int GetNextValue()
        {
            return ++iCount;
        }

    }
}
