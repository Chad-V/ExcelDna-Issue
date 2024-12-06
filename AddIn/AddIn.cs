using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Extensions.DependencyInjection;

namespace MyAddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoClose()
        {

        }

        public void AutoOpen()
        {
            // Create CTP on open
            var pane = CustomTaskPaneManager.CreateCustomTaskPane(new TestingPane(), "DEMO.");

            pane.Show();
        }
    }
}