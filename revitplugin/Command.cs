using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;


namespace RevitPlugin
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class Command : IExternalCommand
    {
        Autodesk.Revit.UI.UIApplication m_Revit;
        
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData,
            ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            Transaction transaction = new Transaction(commandData.Application.ActiveUIDocument.Document, "External Tool");
            m_Revit = commandData.Application;
           
            try
            {
                transaction.Start();
                Autodesk.Revit.ApplicationServices.Application aPPlication = commandData.Application.Application;
                Autodesk.Revit.UI.UIApplication reVit = commandData.Application;
                Document dOcument = commandData.Application.ActiveUIDocument.Document;
                Autodesk.Revit.DB.View view;
                view = dOcument.ActiveView;

                # region FindingFloorslabs

                //Filter out the roofbase elements (roofs slabs)
                FilteredElementIterator iter = (new FilteredElementCollector(commandData.Application.ActiveUIDocument.Document))
                    .OfClass(typeof(RoofBase)).GetElementIterator();
                iter.Reset();

                while (iter.MoveNext())
                {
                    Object obj = iter.Current;

                    //Check whether the object is slab
                    if (obj is RoofBase)
                    {
                        RoofBase slabFloor = obj as RoofBase;
                        string level, elementID;
                        slabFloor = (RoofBase)iter.Current;
                        level = slabFloor.Level.Name;
                        elementID = slabFloor.Id.IntegerValue.ToString();

                        MessageBox.Show("Select the Walls and Railings along the boundary of the Slab at " + level);

                        IList<Reference> reFs = null;
                        List<ElementId> ids = null;

                        try
                        {
                            Selection seLect = commandData.Application.ActiveUIDocument.Selection;
                            reFs = seLect.PickObjects(ObjectType.Element, "Select the Walls and Railings along the boundary of the Slab");
                            ids = new List<ElementId>(reFs.Select<Reference, ElementId>(r => r.ElementId));
                            reFs = null;

                            object oBJect;
                            foreach (ElementId iD in ids)
                            {
                                oBJect = (object)dOcument.GetElement(iD);

                                //Calls shared parameter routines

                                bool successAddParameter = SetNewParameter();
                                SetValueElementInfoParameter(oBJect, elementID, "Slab");

                                //end of code to call shared parameter routines                              

                            }

                        }
                        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                        {
                            return Result.Cancelled;
                        }

                    }
                } 

                # endregion FindingFloorslabs


                # region FindingOpenings

                //Filterout the openings       

                iter = (new FilteredElementCollector(commandData.Application.ActiveUIDocument.Document))
                    .OfClass(typeof(Opening)).GetElementIterator();
                iter.Reset();

                // loop to iterate through elements
                while (iter.MoveNext())
                {
                    Object obj = iter.Current;

                    //Check whether the object is opening
                    if (obj is Opening)
                    {
                        Opening opening = obj as Opening;
                        string level, elementID;
                        opening = (Opening)iter.Current;

                        //find the size of opening
                        BoundingBoxXYZ boundingBox = opening.get_BoundingBox(view);

                        XYZ min, max;
                        min = boundingBox.Min;
                        max = boundingBox.Max;
                        float o_length;

                        o_length = (float)Math.Sqrt(Math.Pow((max.X - min.X), 2) + Math.Pow((max.Y - min.Y), 2) + Math.Pow((max.Z - min.Z), 2));
                        o_length *= (float)0.3048;

                        //check for opening dimension
                        if (o_length >= 0.056)
                        {
                            if (opening.Host.Category.Name.ToString() == "Roofs")
                            {
                                level = opening.Host.Level.Name;
                                elementID = opening.Host.Id.IntegerValue.ToString();
                                MessageBox.Show("Select the Walls and Railings along the boundary of the Opening at " + level);

                                IList<Reference> reFs = null;
                                List<ElementId> ids = null;

                                try
                                {
                                    Selection seLect = commandData.Application.ActiveUIDocument.Selection;
                                    reFs = seLect.PickObjects(ObjectType.Element, "Select the Walls and Railings along the boundary of the Opening");
                                    ids = new List<ElementId>(reFs.Select<Reference, ElementId>(r => r.ElementId));
                                    reFs = null;

                                    object oBJect;
                                    foreach (ElementId iD in ids)
                                    {
                                        oBJect = (object)dOcument.GetElement(iD);

                                        //Calls shared parameter routines

                                        bool successAddParameter = SetNewParameter();
                                        SetValueElementInfoParameter(oBJect, elementID, "Opening");

                                        //end of code to call shared parameter routines                              

                                    }


                                }
                                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                                {
                                    return Result.Cancelled;
                                }
                            }
                        }

                    }
                } 
                # endregion FindingOpenings

                 
                //Filter out the windows
                ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(FamilyInstance));
                ElementCategoryFilter windowsCategoryfilter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
                
                # region FindingWindows
                LogicalAndFilter windowsInstancesFilter = new LogicalAndFilter(familyInstanceFilter, windowsCategoryfilter);
                iter = new FilteredElementCollector(dOcument).WherePasses(windowsInstancesFilter).GetElementIterator();

                double height, width;

                iter.Reset();
                while (iter.MoveNext())
                {
                    FamilyInstance window = iter.Current as FamilyInstance;
                    
                    string elementID;


                    height = window.Symbol.get_Parameter(BuiltInParameter.WINDOW_HEIGHT).AsDouble();
                    width=window.Symbol.get_Parameter(BuiltInParameter.WINDOW_WIDTH).AsDouble();

                    height *= (float)0.3048;
                    width *= (float)0.3048;
                    
                    //check for window opening size
                    if(height>=width && height>=0.48)
                    {
                        elementID = window.Host.Id.IntegerValue.ToString();
                        object oBJect;
                        oBJect = (object)window;

                        bool successAddParameter = SetNewParameter();
                        SetValueElementInfoParameter(oBJect, elementID, "");
                    }
                    else if (height <= width && width>=0.48)
                    {
                        elementID = window.Host.Id.IntegerValue.ToString();
                        object oBJect;
                        oBJect = (object)window;

                        bool successAddParameter = SetNewParameter();
                        SetValueElementInfoParameter(oBJect, elementID, "");
                    }

                } 

                {
                    MessageBox.Show("Select the Windows where there is no Possiblity of Falls");
                    IList<Reference> reFs = null;
                    List<ElementId> ids = null;

                    try
                    {
                        Selection seLect = commandData.Application.ActiveUIDocument.Selection;
                        reFs = seLect.PickObjects(ObjectType.Element, "Select the Windows where there is no Possiblity of Falls");
                        ids = new List<ElementId>(reFs.Select<Reference, ElementId>(r => r.ElementId));
                        reFs = null;

                        object oBJect;
                        foreach (ElementId iD in ids)
                        {
                            oBJect = (object)dOcument.GetElement(iD);

                            //Calls shared parameter routines

                            bool successAddParameter = SetNewParameter();
                            SetValueElementInfoParameter(oBJect, "", "");

                            //end of code to call shared parameter routines                              

                        }


                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }
                }
                # endregion FindingWindows


                //Filter out the stairs
                # region FindingStairs
                try
                {
                    ElementId[] stairIds = new FilteredElementCollector(dOcument).OfCategory(BuiltInCategory.OST_Stairs)
                        .WhereElementIsNotElementType().Select<Element, ElementId>(e => e.Id).ToArray<ElementId>();
                    IList<Reference> reFs = null;
                    List<ElementId> ids = null; 

                    foreach (ElementId iD in stairIds)
                    {
                        foreach (Parameter param in dOcument.GetElement(iD).Parameters)
                        {
                            if (param.Definition.Name.ToString() == "Top Level")
                            {                                
                                string level=dOcument.GetElement(param.AsElementId()).Name;
                                MessageBox.Show("Select the railings and walls along stairs where fall form height is possible at top " + level);

                                Selection seLect = commandData.Application.ActiveUIDocument.Selection;
                                reFs = seLect.PickObjects(ObjectType.Element, "Select railings and stairs where there is Possiblity of Falls");
                                ids = new List<ElementId>(reFs.Select<Reference, ElementId>(r => r.ElementId));
                                reFs = null;

                                object oBJect;
                                foreach (ElementId eleiD in ids)
                                {
                                    oBJect = (object)dOcument.GetElement(eleiD);

                                    //Calls shared parameter routines

                                    bool successAddParameter = SetNewParameter();
                                    SetValueElementInfoParameter(oBJect, iD.IntegerValue.ToString(), "Stair");

                                    //end of code to call shared parameter routines                              

                                }
                            }
                        }

                    }

                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                # endregion FindingStairs


            }
            catch (Exception e)
            {
                message = e.ToString();
                return Autodesk.Revit.UI.Result.Failed;
            }
            finally
            {
                transaction.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        # region Shared parameter
        

        //FUNCTIONS FOR CREATING SHARED PARAMETER

        //create the shared parameter file if not existiinng
        private DefinitionFile AccessOrCreateExternalSharedParameterFile()
        {           

            // The path of ourselves shared parameters file (Shared parameter file is in debug folder)
            string sharedParameterFile = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            sharedParameterFile = sharedParameterFile + "\\SafetySharedParameters.txt";

            //Method's return
            DefinitionFile informationFile = null;

            // Check if the file is exit
            System.IO.FileInfo documentMessage = new System.IO.FileInfo(sharedParameterFile);
            bool fileExist = documentMessage.Exists;

            // Create file for external shared parameter since it does not exist
            if (!fileExist)
            {
                FileStream fileFlow = File.Create(sharedParameterFile);
                fileFlow.Close();
            }

            // Set  ourselves file to  the externalSharedParameterFile 
            m_Revit.Application.SharedParametersFilename = sharedParameterFile;
            informationFile = m_Revit.Application.OpenSharedParameterFile();
            return informationFile;
        }



        public bool SetNewParameter()
        {
            //Open the shared parameters file 
            // via the private method AccessOrCreateExternalSharedParameterFile
            DefinitionFile informationFile = AccessOrCreateExternalSharedParameterFile();

            if (null == informationFile)
            {
                return false;
            }

            // Access an existing or create a new group in the shared parameters file
            DefinitionGroups informationCollections = informationFile.Groups;
            DefinitionGroup informationCollection = null;

            informationCollection = informationCollections.get_Item("MyParameters");

            if (null == informationCollection)
            {
                informationCollections.Create("MyParameters");
                informationCollection = informationCollections.get_Item("MyParameters");
            }

            // Access an existing or create a new external parameter definition 
            // belongs to a specific group
            
            Definition information1 = informationCollection.Definitions.get_Item("ElementInfo");

            if (null == information1)
            {
                informationCollection.Definitions.Create("ElementInfo", Autodesk.Revit.DB.ParameterType.Text);
                information1 = informationCollection.Definitions.get_Item("ElementInfo");
            }

            Definition information2 = informationCollection.Definitions.get_Item("ProtectionInfo");

            if (null == information2)
            {
                informationCollection.Definitions.Create("ProtectionInfo", Autodesk.Revit.DB.ParameterType.Text);
                information2 = informationCollection.Definitions.get_Item("ProtectionInfo");
            }
            

            // Create a new Binding object with the categories to which the parameter will be bound
            CategorySet categories = m_Revit.Application.Create.NewCategorySet();            
            Category CatWall=null;
            Category CatWindow = null;
            Category CatRailing = null;
       
            // use category in instead of the string name to get category 
            CatWall = m_Revit.ActiveUIDocument.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Walls);
            CatWindow = m_Revit.ActiveUIDocument.Document.Settings.Categories.get_Item(BuiltInCategory.OST_Windows);
            CatRailing = m_Revit.ActiveUIDocument.Document.Settings.Categories.get_Item("Railings");
            categories.Insert(CatWall);
            categories.Insert(CatWindow);
            categories.Insert(CatRailing);
            
            InstanceBinding caseTying = m_Revit.Application.Create.NewInstanceBinding(categories);
            
            // Add the binding and definition to the document
            if (m_Revit.ActiveUIDocument.Document.ParameterBindings.Insert(information1, caseTying) 
                && m_Revit.ActiveUIDocument.Document.ParameterBindings.Insert(information2, caseTying))
            {               
                    return true;
            }

            return false; ;
        }



        public void SetValueElementInfoParameter(object localObject, string ElementID, string SalbOrOpeningorStair)
        {
            Element elem;         
            if (localObject is Wall)
            {
                elem = (Wall)localObject; 
                // Find the parameter which is named "ElementInfo"  
                ParameterSet attributes = elem.Parameters;
                System.Collections.IEnumerator iter = attributes.GetEnumerator();
                iter.Reset();
                while (iter.MoveNext())
                {
                    Parameter attribute = iter.Current as Autodesk.Revit.DB.Parameter;
                    Definition information = attribute.Definition;
                    //set the value to the parameter
                    if ((null != information) && ("ElementInfo" == information.Name) )
                    {                        
                        attribute.Set(ElementID);
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name) && (SalbOrOpeningorStair == "Slab"))
                    {
                        attribute.Set("ExteriorWall");
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name) && (SalbOrOpeningorStair == "Opening"))
                    {
                        attribute.Set("Opening");
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name) && (SalbOrOpeningorStair == "Stair"))
                    {
                        attribute.Set("Stair");
                    }
                }
            }

            elem = (Element)localObject;
                
            if (elem.Category.Name.ToString()==m_Revit.ActiveUIDocument.Document
                .Settings.Categories.get_Item(BuiltInCategory.OST_Windows).Name.ToString())
            {                
                // Find the parameter which is named "ElementInfo"                  
                ParameterSet attributes = elem.Parameters;
                System.Collections.IEnumerator iter = attributes.GetEnumerator();
                iter.Reset();
                while (iter.MoveNext())
                {
                    Parameter attribute = iter.Current as Autodesk.Revit.DB.Parameter;
                    Definition information = attribute.Definition;
                    //set the value to the parameter
                    if ("ElementInfo" == information.Name)
                    {
                        attribute.Set(ElementID);
                    }                    
                }
            }

            if (elem.Category.Name.ToString() == "Railings")
            {
                ParameterSet attributes = elem.Parameters;
                System.Collections.IEnumerator iter = attributes.GetEnumerator();
                iter.Reset();
                while (iter.MoveNext())
                {
                    Parameter attribute = iter.Current as Autodesk.Revit.DB.Parameter;
                    Definition information = attribute.Definition;
                    //set the value to the parameter
                    if ((null != information) && ("ElementInfo" == information.Name) )
                    {
                        attribute.Set(ElementID);
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name) && (SalbOrOpeningorStair == "Slab"))
                    {
                        attribute.Set("ExteriorWall");
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name)&& (SalbOrOpeningorStair == "Opening"))
                    {
                        attribute.Set("Opening");
                    }
                    if ((null != information) && ("ProtectionInfo" == information.Name) && (SalbOrOpeningorStair == "Stair"))
                    {
                        attribute.Set("Stair");
                    }
                }
            }
        }
        # endregion Sahered Parameter
    }
}
