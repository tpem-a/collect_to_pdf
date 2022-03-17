using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace project
{
    public partial class SolidWorksMacro
    {

        ModelDoc2 swModel;
        SelectionMgr swSelMgr;
        Feature swFeat;
        Feature swSubFeat;
        ModelDocExtension swExt;

        bool bRet;
        int lRet;

        string save_path;

        List<string> begin_msg = new List<string>() { };
        List<string> draws = new List<string>() { };
        List<string> no_draws = new List<string>() { };
        List<string> end_msg = new List<string>() { };

        public void TraverseFeatureFeatures(Feature swFeat, long nLevel)
        {
            ///Feature swSubFeat;
            Feature swSubSubFeat;
            Feature swSubSubSubFeat;
            string sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel; i++)
            {
                sPadStr = sPadStr + " ";
            }
            while ((swFeat != null))
            {
                Debug.Print(sPadStr + swFeat.Name + " [" + swFeat.GetTypeName2() + "]");
                swSubFeat = (Feature)swFeat.GetFirstSubFeature();

                while ((swSubFeat != null))
                {
                    Debug.Print(sPadStr + "  " + swSubFeat.Name + " [" + swSubFeat.GetTypeName() + "]");
                    swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                    while ((swSubSubFeat != null))
                    {
                        Debug.Print(sPadStr + "    " + swSubSubFeat.Name + " [" + swSubSubFeat.GetTypeName() + "]");
                        swSubSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                        while ((swSubSubSubFeat != null))
                        {
                            Debug.Print(sPadStr + "      " + swSubSubSubFeat.Name + " [" + swSubSubSubFeat.GetTypeName() + "]");
                            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();

                        }

                        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();

                    }

                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();

                }

                swFeat = (Feature)swFeat.GetNextFeature();

            }

        }

        public void TraverseComponentFeatures(Component2 swComp, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swComp.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }

        public void TraverseComponent(Component2 swComp, long nLevel)
        {
            object[] vChildComp;
            Component2 swChildComp;
            string sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel - 1; i++)
            {
                sPadStr = sPadStr + " ";
            }

            vChildComp = (object[])swComp.GetChildren();
            for (i = 0; i < vChildComp.Length; i++)
            {
                swChildComp = (Component2)vChildComp[i];
                

                array_drw(swChildComp.GetPathName());

                ///TraverseComponentFeatures(swChildComp, nLevel);
                TraverseComponent(swChildComp, nLevel + 1);
            }
        }

        public void TraverseModelFeatures(ModelDoc2 swModel, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swModel.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }
        
        public void array_drw(string pth)
        {
            if (!draws.Contains(pth))
            {
                draws.Add(pth.Remove(pth.Length - 3, 3) + "DRW");
                bool i = exist_drw(pth);

                    if (!i)
                    {

                        no_draws.Add(pth);

                    }
                    else
                    {

                        save_pdf(pth, save_path);
                        ///Debug.Print("- Save pdf: " + pth);

                    }

                

                //Debug.Print("   add new   " + i);
            }
        }

        public bool exist_drw(string ptch)
        {

            ptch = ptch.Remove(ptch.Length - 3, 3) + "DRW";

            var exist = File.Exists(ptch);
            
            return exist;
        }

        public void save_pdf(string pt, string fold)
        {

            ///Debug.Print("-- " + pt + " / " + fold);


            ExportPdfData swExportPDFData = default(ExportPdfData);
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModExt = default(ModelDocExtension);
            DrawingDoc swDrawDoc = default(DrawingDoc);
            Sheet swSheet1 = default(Sheet);
            bool boolstatus1 = false;
            string filename1 = null;
            int errors = 0;
            int warnings = 0;
            string[] obj = null;

            swModel = (ModelDoc2)swApp.OpenDoc6(pt.Remove(pt.Length - 3, 3) + "DRW", (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            string ss = swModel.GetTitle();
            Debug.Print("-" + ss);

            swModExt = (ModelDocExtension)swModel.Extension;
            swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

            // Get the names of the drawing sheets in the drawing
            // to get the size of the array of drawing sheets
            swDrawDoc = (DrawingDoc)swModel;

            obj = (string[])swDrawDoc.GetSheetNames();
            int count = 0;
            count = obj.Length;
            int i = 0;
            object[] objs = new object[count - 1];
            DispatchWrapper[] arrObjIn = new DispatchWrapper[count - 1];

            for (i = 0; i < count - 1; i++)
            {
                boolstatus1 = swDrawDoc.ActivateSheet((obj[i]));
                swSheet1 = (Sheet)swDrawDoc.GetCurrentSheet();
                objs[i] = swSheet1;
                arrObjIn[i] = new DispatchWrapper(objs[i]);
            }

            string nm = pt.Split('\\').Last();
            

            // Save the drawings sheets to a PDF file
            boolstatus1 = swExportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (arrObjIn));
            
            ///открыть после сохранения
            swExportPDFData.ViewPdfAfterSaving = false;

            //string folderName = fold + '\\' + nm.Remove(nm.Length - 7, 7) + ".pdf";

            //создаём новую папку
            if (!Directory.Exists(fold))
            {
                Directory.CreateDirectory(fold);
                //Debug.Print("create folder!");
                begin_msg.Add("Create folder");
            }
            

                boolstatus1 = swModExt.SaveAs(fold+ "\\" + nm.Remove(nm.Length - 7, 7) + ".pdf", (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings);

            ///Debug.Print("* " + fold);

            filename1 = swModel.GetTitle();
            swApp.QuitDoc(filename1);

            //Debug.Print("- Save pdf: " + pt + " / " + fold);
        }

        public void Main()
        {

            ModelDoc2 swModel;
            ConfigurationManager swConfMgr;
            Configuration swConf;
            Component2 swRootComp;

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
            swConf = (Configuration)swConfMgr.ActiveConfiguration;
            swRootComp = (Component2)swConf.GetRootComponent();

            

            System.Diagnostics.Stopwatch myStopwatch = new Stopwatch();
            myStopwatch.Start();

            //Debug.Print("File = " + swModel.GetPathName());
            begin_msg.Add("Project open: " + swModel.GetPathName());
            
            draws.Add(swModel.GetPathName());

            string pp = swModel.GetPathName();

            int position = pp.LastIndexOf('\\');

            ///Debug.Print(swModel.GetPathName().Remove(swModel.GetPathName().Length - 7, 7));
            save_path = swModel.GetPathName().Remove(swModel.GetPathName().Length - 7, 7);

            save_pdf(swModel.GetPathName(), save_path);

            

            //Form frm = new start();
            ///frm.ShowDialog();

            
            ///TraverseModelFeatures(swModel, 1);

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                TraverseComponent(swRootComp, 2);
            }


            foreach (var dr in no_draws)
            {
                Debug.Print("---" + dr);
            }

            foreach (var dr2 in draws)
            {
                Debug.Print("+++" + dr2);
            }

            Debug.Print("Уникальных деталей: "+draws.Count.ToString());
            Debug.Print("Нет чертежей на: " + no_draws.Count.ToString() + " шт.");


            myStopwatch.Stop();
            TimeSpan myTimespan = myStopwatch.Elapsed;
            Debug.Print("Time = " + myTimespan.TotalSeconds + " sec");

            ///MessageBox.Show("Нет чертежей на: " + no_draws.Count.ToString() + " шт.", "Команда выполнена за" + Math.Round(myTimespan.TotalSeconds) +" c");

            
            int tm = (int)myTimespan.TotalSeconds;

            end_msg.Add("Жыве Беларусь! Слава Украине!");

            Form frm = new start(tm,begin_msg,draws,no_draws,end_msg);
            frm.ShowDialog();
            

        }


        /// <summary> 
        /// The SliaWorks swApp variable is pre-assigned for you. 
        /// </summary> 
        public SldWorks swApp;
    }


}

