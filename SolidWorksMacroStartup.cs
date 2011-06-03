using SolidWorks.Interop.sldworks;
using System;

namespace IGESExporter.csproj
{
    public partial class SolidWorksMacro
    {
        private void SolidWorksMacro_Startup(object sender, EventArgs e)
        {
            this.swApp = (SldWorks)this.SolidWorksApplication;
            SolidWorks.RunMacroResult runResult = RunMacro();

            if (runResult == SolidWorks.RunMacroResult.NoneSpecified)
            {
                try
                {
                    SwMacroSetup();
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex.ToString()); }
                Main();
                try
                {
                    SwMacroCleanup();
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex.ToString()); }
            }
        }

        private void SolidWorksMacro_Shutdown(object sender, EventArgs e)
        {

        }


        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(SolidWorksMacro_Startup);
            this.Shutdown += new System.EventHandler(SolidWorksMacro_Shutdown);
        }


    }
}


