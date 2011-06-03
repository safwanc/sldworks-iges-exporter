using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms; 
using System;
using System.IO;
using System.Diagnostics; 

namespace IGESExporter.csproj
{
    public partial class SolidWorksMacro
    {
        public SldWorks swApp;
        public bool swSuccess;
        private int errors;
        private int warnings;
        private Component2 swComponent; 

        public void Main()
        {
            swSuccess = false; errors = 0; warnings = 0; 
            string swSourcePath = swApp.GetCurrentWorkingDirectory();
            string swExportPath = Path.Combine(swSourcePath, "IGS");
            string[] tmp = Directory.GetFiles(swSourcePath, "*.SLDPRT", SearchOption.TopDirectoryOnly);

            if (Directory.Exists(swExportPath))
                Directory.Delete(swExportPath, true);

            Directory.CreateDirectory(swExportPath);
            ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
            
            AssemblyDoc swAssem = (AssemblyDoc)swModel;

            #region TopLevel
            string igsFileName = MakeFileName(swModel.Extension.Document.GetTitle()); 
            swModel.Extension.SaveAs(Path.Combine(swExportPath, igsFileName), (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Copy, null, ref errors, ref warnings);
            //PrintSaveResults();             
            #endregion TopLevel

            #region Subassemblies
            string swExportSubPath = Path.Combine(swExportPath, "Subassemblies"); 
            Directory.CreateDirectory(swExportSubPath); 
            object[] objComps = (object[])swAssem.GetComponents(true);
            ModelDoc2 igsModel; 

            foreach (object obj in objComps)
            {
                swComponent = obj as Component2;
                Debug.WriteLine(String.Format("Working on {0}.. \n\t {1}", swComponent.Name2, swComponent.GetPathName())); 
                igsModel = swApp.OpenDoc6(swComponent.GetPathName(), (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly, null, ref errors, ref warnings);
                swApp.ActivateDoc2(igsModel.GetTitle(), true, ref errors); 
                igsFileName = MakeFileName(igsModel.Extension.Document.GetTitle());
                igsModel.Extension.SaveAs(Path.Combine(swExportSubPath, igsFileName), (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Copy, null, ref errors, ref warnings);
                swApp.CloseDoc(igsModel.GetTitle()); 
            }
            #endregion Subassemblies

            #region Parts
            string swExportPartPath = Path.Combine(swExportPath, "Parts");
            Directory.CreateDirectory(swExportPartPath); 
            string[] swPartFiles = Directory.GetFiles(swSourcePath, "*.SLDPRT", SearchOption.TopDirectoryOnly);
            foreach (string partFile in swPartFiles)
            {
                try
                {
                    igsModel = swApp.OpenDoc6(partFile, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly, null, ref errors, ref warnings);
                    swApp.ActivateDoc2(igsModel.GetTitle(), true, ref errors);
                    string[] swConfigs = (string[])igsModel.GetConfigurationNames();
                    int swConfigCount = igsModel.GetConfigurationCount();
                    string swConfigFilename; 

                    for (int i = 0; i < swConfigCount; i++)
                    {
                        igsModel.ShowConfiguration2(swConfigs[i]);
                        swConfigFilename = igsModel.Extension.Document.GetTitle(); 
                        if (swConfigCount > 1)
                            swConfigFilename += String.Format("-CFG{0}", (i+1).ToString());
                        igsFileName = MakeFileName(swConfigFilename);
                        igsModel.Extension.SaveAs(Path.Combine(swExportPartPath, igsFileName), (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Copy, null, ref errors, ref warnings);
                    }
                    swApp.CloseDoc(igsModel.GetTitle());
                }
                catch (NullReferenceException e)
                {
                    Debug.WriteLine(e.Message); 
                }
            }
            #endregion Parts

            string timeFilename = Path.Combine(swExportPath, "Timestamp.txt");

            using (StreamWriter swTimeStamper = new StreamWriter(timeFilename))
            {
                swTimeStamper.WriteLine("Exported on {0} at {1}", System.DateTime.Now.ToShortDateString(), System.DateTime.Now.ToShortTimeString());
                swTimeStamper.Flush(); 
            }
        }

        public void ShowBox(string t)
        {
            swApp.SendMsgToUser2(t, (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk); 
        }

        private string MakeFileName(string title)
        {
            return title + ".IGS"; 
        }

        private void PrintSaveResults()
        {
            if (errors > 0)
            {
                Debug.WriteLine(Enum.GetName(typeof(swFileSaveError_e), (swFileSaveError_e)errors));
                errors = 0; 
            }
            
            if (warnings > 0)
            {
                Debug.WriteLine(Enum.GetName(typeof(swFileSaveWarning_e), (swFileSaveWarning_e)warnings));
                warnings = 0; 
            }
        }



    }
}