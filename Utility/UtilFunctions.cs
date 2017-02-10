using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Uniconta.Common;

namespace ImportingTool.Utility
{
    public class UtilFunctions
    {

        public static void ShowErrorMessage(ErrorCodes ec)
        {
            MessageBox.Show(ec.ToString());
        }
    }
}
