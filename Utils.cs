using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookToMariadb
{
    internal static class Utils
    {
        public static string ToBase64String(string s)
        {
            string res = string.IsNullOrEmpty(s) ? "" : s;
            return System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(res));
        }
    }
}
