using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1
{
    public class ExcelInstCollector
    {
        private List<string> insts = null;

        public ExcelInstCollector()
        {
            insts = new List<string>();
            foreach (var item in Process.GetProcessesByName("excel"))
            {
                insts.Add(item.ProcessName.ToString());
            }
            
        }
        public List<string> ActiveInstances
        {
            get
            {
                return insts;
            }
        }
    }
    
}
