using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

//https://limbioliong.wordpress.com/2011/07/14/passing-a-safearray-of-managed-structures-to-unmanaged-code/

namespace COM_blog_lab
{
    class Program
    {
        [ComVisible(true)]
        [Guid("B4A16864-42FF-48ea-973B-E0BE5922719E")]
        public struct TestStruct
        {
            [MarshalAs(UnmanagedType.I4)]
            public Int32 m_integer;
            [MarshalAs(UnmanagedType.R8)]
            public double m_double;
            [MarshalAs(UnmanagedType.BStr)]
            public string m_string;
        }

        [ComVisible(true)]
        [Guid("FE37FC87-1E1B-4262-B720-1E6503B3E964")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ITestInterface
        {
            void GetTestStructArray([Out] out TestStruct[] test_struct_array);
            void SetTestStructArray([In] TestStruct[] test_struct_array);
            void ReferenceTestStructArray([In][Out] ref TestStruct[] test_struct_array);
        }
        static void Main(string[] args)
        {
        }
    }
}
