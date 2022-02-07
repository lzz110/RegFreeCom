using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TestCom
{
    

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("32968C43-E1F4-4AF2-921C-50A2FB63FA2E")]
    [ProgId("TestCom.MyCom123")]
    public class Class1
    {
        private static int iCount = 100;

        [ComVisible(true)]
        public int GetLastValue()
        {
            return --iCount;
        }

    }
}
