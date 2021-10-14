using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelAddIn1.Utils;
using ExcelAddIn1.Forms;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private Excel.Application _app;

        FileLoader fileLoader = new FileLoader();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _app = Globals.ThisAddIn.Application;

        }

        private void bntImport_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result = fileLoader.OpenFileDialogForm();
            if (result != DialogResult.OK) return;

            FormLoading loader = new FormLoading();
            loader.StartPosition = FormStartPosition.CenterScreen;
            loader.Show();
            // This is a pretty important part because on the FromCurrentSynchronizationContext() call you can get Exception
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            }

            var context = TaskScheduler.FromCurrentSynchronizationContext();
            var task = new Task(() =>
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                app.Interactive = false;
                fileLoader.loadLua(app);
                //fileLoader.loadDataToMem(app);

                // This code is used to call Close() method Thread safe otherwise you get Cross-thread operation not valid Exception
                ThreadSafeHelper.InvokeControlMethodThreadSafe(loader, () =>
                {
                    loader.Close();
                });

                app.StatusBar = "Ready";
                app.Interactive = true;
            });

            task.Start();

        }

        private void bntExport_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Export");
        }

        private void bntVerify_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Export");
        }
    }
}
