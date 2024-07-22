using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace ExportFlatPatternsInPartToDxf
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {

            //Extension can either be .dxf or .dwg
            string exportExtension = ".dxf";

            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;

            //Make sure that the open document is a part
            if (swDoc.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                MessageBox.Show("This is not a part document. Please activate a part document and try again", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int uniqueSheetMetalBodiesCount = 0;
            int cutListFolderNumber = 0;
            Feature swFeat;
            FeatureManager featMgr = (FeatureManager)swDoc.FeatureManager;
            object[] vFeatures = (object[])featMgr.GetFeatures(true);
            int featCount = featMgr.GetFeatureCount(true);

            for (int i = 0; i < featCount; i++)
            {
                //This loop is just counting how many cutlist folders exist in a single part document
                //If more than one cutlist folder exits, it means there are multiple flat patterns that need to be exported.
                //This requires adding -1 -2 -3 etc.. to the dxf or dwg name when exporting the flat patterens
                //This logic is implemented below starting at --> if (uniqueSheetMetalBodiesCount > 1)..
                swFeat = (Feature)vFeatures[i];

                if (swFeat.GetTypeName2().Contains("CutListFolder")) uniqueSheetMetalBodiesCount++;
            }

            for (int i = 0; i < featCount; i++)
            {
                swFeat = (Feature)vFeatures[i];

                if (swFeat.GetTypeName2().Contains("CutListFolder"))
                { 
                    cutListFolderNumber++;
                    BodyFolder BodyFolder = swFeat.GetSpecificFeature2() as BodyFolder;
                    BodyFolder.UpdateCutList();

                    int BodyCount = BodyFolder.GetBodyCount();

                    if (BodyCount < 1) return;

                    object[] vBodies = (object[])BodyFolder.GetBodies();

                    if (vBodies == null) return;

                    //Only taking the first body in the cut list folder since a cutt list folder can have multiple identical bodies if the bodies were patterned or mirrored
                    Body2 Body = (Body2)vBodies[0];
                    object[] bodyFeatures = (object[])Body.GetFeatures();
                    Feature swbodyFeat;

                    foreach (var bodyFeat in bodyFeatures)
                    {
                        swbodyFeat = (Feature)bodyFeat;

                        if (swbodyFeat.GetTypeName2().Contains("FlatPattern"))
                        {
                            swbodyFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, 1, null);
                            swbodyFeat.Select2(false, 0);

                            PartDoc swPart = (PartDoc)swDoc;
                            string modelPath = swDoc.GetPathName();
                            string directory = Path.GetDirectoryName(modelPath);
                            string modelName = Path.GetFileNameWithoutExtension(modelPath);
                            string dxfFileName;

                            if (uniqueSheetMetalBodiesCount > 1)
                            {
                                dxfFileName = modelName + "-" + cutListFolderNumber.ToString() + exportExtension;
                            }
                            else
                            {
                                //if the is only one cut list folder then no need to add -1 to the dxf name because the part contains only one flat pattern
                                dxfFileName = modelName + exportExtension;
                            }

                            string dxfFilePath = Path.Combine(directory, dxfFileName);

                            //sheetMetalExportOptions is set to 1 to export only the flat pattern geometry  

                            int sheetMetalExportOptions = 1;

                            swPart.ExportToDWG2(dxfFileName, modelPath,
                                                        (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true,
                                                         null, false, false, sheetMetalExportOptions, null);

                            swbodyFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, 1, null);
                        }
                    }
                }
            }
            return;
        }

        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;

    }
}

