using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.P3dProjectParts;
using Autodesk.ProcessPower.PartsRepository;
//using Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;
using Autodesk.ProcessPower.PnP3dDataLinks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.P3dUI;


// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF 
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE 
// UNINTERRUPTED OR ERROR FREE. 
//V1
//added autotray
//V2
//include shortdescription as selection criteria
//extract dimenensions string, replace just length (must be the L parameter!), W must be set up in the spec as nominal diameter (size)
//can do imperial with imperial spec and metric with metric spec
//handle just parts from the specified spec
//use size and spec from the existing pipes not from the ribbon, exception: "setSpecForOneTray" uses the size from the ribbon
//replace the spec string in the properties for the couplings with the empty string, so they cannot be changed by the spec update
//v3
//can now flip elbows of any angle
//use top view for the autotray command, it will connect (general: view from the "up direction" that you will choose, then it will connect)!
//length of coupling will be displayed in the length field, it will show the length at the time when the autotray command is executed
//H of coupling will be written to the DesignStd field, it can then be used by a PLANTDEFINECALCPROPERTIES for correct BOP calculation: BOP+MatchingPipeOd/2-ToNumber(DesignStd)/2
//v4
//removed the H property writing, because it is easier to correct the BOP by editing the Outerdiamter OD of the cable trays
//now assigning the linenumbertag from the pipe to the coupling (tray)
//disabled setSpecForOneTray command
//added 2 pick for direction definition
//added respect to the current user coordinate system 
//v5 test with sleeve
//v6 going back to coupling, not sent out yet

//todo: still the ribbon setting must fit the pipes ND, because setSpecForOneTray is part of the autotray command, this can be fixed
[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(null)]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(pssCommands.Program))]

namespace pssCommands
{
    class Program
    {

        public static SpecPart FetchSpecPart(
                  String PartName,
                  string shortdesc,
                  string SpecName,
                  NominalDiameter NomDiameter,
                  ref Editor ed)
        {
            var specMgr = SpecManager.GetSpecManager();

            SpecPart specPart = null;
            if (specMgr.HasType(SpecName, PartName))
            {
                SpecPartReader specPartReader = specMgr.SelectParts(SpecName, PartName);

                while (specPartReader.Next())
                {
                    specPart = specPartReader.Current;

                    if (specPart.Type.Equals(PartName) &&
                        shortdesc.Equals(specPart.PropValue("ShortDescription").ToString()) &&
                        NomDiameter.Value.ToString().Equals(specPart.PropValue("Size").ToString().Replace("\"", "")))
                    {
                        break;
                    }
                }

            }

            return specPart;
        }

        public static PipeInlineAsset CreateInlineAsset()
        {
            PipeInlineAsset pipeInlineAssetPart = new PipeInlineAsset();
            return pipeInlineAssetPart;
        }


        public static ObjectId AddObjectToDatabase(
            SpecPart specPart,
            Autodesk.ProcessPower.PnP3dObjects.Part part,
            String connectionName,
            PartSizePropertiesCollection connectionPropColl,
            ref Database db,
            ref Editor ed,
            ref Project currentProject,
            ref DataLinksManager dlm,
            ref DataLinksManager3d dlm3d,
            ref PipingObjectAdder pipeObjAdder)
        {
            ObjectId partObjectId = ObjectId.Null;
            if (pipeObjAdder == null)
            {
                ed.WriteMessage("Error: Cannot create PipingObjectAdder");
                return partObjectId;
            }


            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {

                    PipeInlineAsset pipeInlineAsset = part as PipeInlineAsset;

                    pipeObjAdder.Add(specPart, pipeInlineAsset);


                    partObjectId = part.ObjectId;
                    trans.AddNewlyCreatedDBObject(part, true);
                    trans.Commit();
                }
            }
            catch (SystemException ex)
            {
                ed.WriteMessage("Error: Exception while appending objects.\n");
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                ed.WriteMessage(trace.ToString());
                ed.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                ed.WriteMessage("message: " + ex.Message);
            }
            return partObjectId;
        }

        [CommandMethod("execTrayCommand", CommandFlags.UsePickSet)]
        public static void execTrayCommand()
        {
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Database db = AcadApp.DocumentManager.MdiActiveDocument.Database;
                Project currentProject = PlantApp.CurrentProject.ProjectParts["Piping"];
                DataLinksManager dlm = currentProject.DataLinksManager;
                string configstr = "";
                string tcom = "";
                string cspec = "";
                string cspecpath = "";
                string shortdesc = "";
                string W = "";
                string H = "";
                NominalDiameter nd = new NominalDiameter();
                UISettings sett = new UISettings();
                nd = NominalDiameter.FromDisplayString(null, sett.CurrentSize);
                //

                PromptResult pr = ed.GetString("\nconfiguration string: ");
                if (pr.Status != PromptStatus.OK)
                {
                }
                else
                    configstr = pr.StringResult;

                if (!configstr.Equals(""))
                {
                    string[] configArr = configstr.Split(new char[] { ',' });

                    foreach (string cstr in configArr)
                    {
                        string cstrkey = cstr.Split(new char[] { '=' })[0].Trim();
                        string cstrval = cstr.Split(new char[] { '=' })[1].Trim();

                        switch (cstrkey)
                        {
                            case "tcom":
                                tcom = cstrval;
                                if (tcom.Equals("")) goto default;
                                break;
                            case "cspec":
                                if (cstrval.Equals(""))
                                {
                                    goto default;
                                }
                                else if (cstrval.IndexOf("/") == -1)
                                {
                                    cspec = cstrval;
                                    PlantProject currentProj = PlantApplication.CurrentProject;
                                    cspecpath = currentProj.ProjectFolderPath + "\\Spec Sheets\\" + cstrval + ".pspx";
                                }
                                else
                                {
                                    cspecpath = cstrval;
                                    cspec = cstrval.Substring(cstrval.LastIndexOf("/") + 1);
                                    cspec = cspec.Substring(0, cspec.Length - 5);
                                }
                                break;
                            case "shortdesc":
                                shortdesc = cstrval;
                                if (shortdesc.Equals("")) goto default;
                                break;
                            default:
                                ed.WriteMessage("\nconfiguration string is not according to the rules!");
                                return;
                                break;
                        }
                    }


                }
                else
                {
                    ed.WriteMessage("No configuration string was provided\n");
                    return;
                }


                if (tcom.Equals("setSpecForOneTray"))
                {
                    /*PromptPointResult SelectedPoint = ed.GetPoint("\nPunkt1: ");
                    Point3d startLoc = SelectedPoint.Value;

                    SelectedPoint = ed.GetPoint("\nPunkt2: ");
                    Point3d endLoc = SelectedPoint.Value;

                    setSpecData(ref ed, startLoc, endLoc, shortdesc, nd, H, cspecpath, ref db);*/
                }
                else
                {
                    TypedValue[] filterlist = new TypedValue[1];
                    filterlist[0] = new TypedValue(0, "ACPPPIPE,ACPPPIPEINLINEASSET");//
                    SelectionFilter filter = new SelectionFilter(filterlist);

                    PromptSelectionResult selRes = ed.GetSelection(filter);
                    if (selRes.Status != PromptStatus.OK)
                    { return; }

                    //prompt for zdir
                    PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                    pKeyOpts.Message = "\nEnter the up direction ";
                    pKeyOpts.Keywords.Add("Up");
                    pKeyOpts.Keywords.Add("Down");
                    pKeyOpts.Keywords.Add("North");
                    pKeyOpts.Keywords.Add("South");
                    pKeyOpts.Keywords.Add("West");
                    pKeyOpts.Keywords.Add("East");
                    pKeyOpts.Keywords.Add("2pick");
                    pKeyOpts.AllowNone = false;

                    string zdirSel = ed.GetKeywords(pKeyOpts).StringResult;

                    Vector3d zVector = new Vector3d();
                    bool pick2 = false; 

                    switch (zdirSel)
                    {
                        case "Up":
                            zVector = new Vector3d(0, 0, 1);
                            break;
                        case "Down":
                            zVector = new Vector3d(0, 0, -1);
                            break;
                        case "North":
                            zVector = new Vector3d(0, 1, 0);
                            break;
                        case "South":
                            zVector = new Vector3d(0, -1, 0);
                            break;
                        case "West":
                            zVector = new Vector3d(-1, 0, 0);
                            break;
                        case "East":
                            zVector = new Vector3d(1, 0, 0);
                            break;
                        case "2pick":
                            pick2 = true;
                            break;
                    }

                    if (pick2)
                    {
                        PromptPointResult SelPoint = ed.GetPoint("\nPunkt1: ");
                        Point3d sLoc = SelPoint.Value;
                        SelPoint = ed.GetPoint("\nPunkt2: ");
                        Point3d eLoc = SelPoint.Value;
                        zVector = sLoc.GetVectorTo(eLoc);
                    }

                    zVector = zVector.TransformBy(ed.CurrentUserCoordinateSystem);

                    ObjectIdCollection objIds = new ObjectIdCollection();
                    ObjectId[] objIdArray = selRes.Value.GetObjectIds();


                    foreach (ObjectId id in objIdArray)
                    {
                        Point3d startLoc = new Point3d();
                        Point3d endLoc = new Point3d();

                        int rowId = dlm.FindAcPpRowId(id);

                        StringCollection pnames = new StringCollection();
                        pnames.Add("PnPClassName");
                        pnames.Add("Spec");
                        pnames.Add("Size");
                        pnames.Add("COG Z"); 
                        //pnames.Add("LineNumberTag");
                        //todo add spec size

                        StringCollection theprops = dlm.GetProperties(rowId, pnames, true);
                        if (!theprops[1].Equals(cspec))
                            continue;

                        if (theprops[0].Equals("Pipe"))
                        {
                            int lgrowId = 0;
                            Pipe entpipe = new Pipe();
                            nd = new NominalDiameter(Convert.ToDouble(theprops[2].Replace("\"", "")));
                            using (Transaction tr = db.TransactionManager.StartTransaction())
                            {
                                entpipe = tr.GetObject(id, OpenMode.ForWrite) as Pipe;
                                startLoc = entpipe.StartPoint;
                                endLoc = entpipe.EndPoint;
                                int prowId = dlm.FindAcPpRowId(id);
                                PnPRowIdArray prow = dlm.GetRelatedRowIds("P3dLineGroupPartRelationship", "Part", prowId, "LineGroup");
                                lgrowId = prow.First.Value;
                                entpipe.Erase();
                                tr.Commit();
                            }

                            H = setSpecData(ref ed, startLoc, endLoc, shortdesc, nd, H, cspecpath, ref db);
                            ed.WriteMessage("in " + nd.Value.ToString() + " and " + cspec);
                            ObjectId coupId = insertSpecPartTo3d(startLoc, endLoc, zVector, ref ed, nd, shortdesc, cspec, ref db, ref currentProject, ref dlm);

                            //string newBoP = (Convert.ToDouble(theprops[3]) - (Convert.ToDouble(H) / 2)).ToString();

                            int crowId = dlm.FindAcPpRowId(coupId);

                            StringCollection cnames = new StringCollection();
                            cnames.Add("Spec");
                            //cnames.Add("LineNumberTag");
                            cnames.Add("Length");
                            StringCollection cvals = new StringCollection();
                            cvals.Add("");
                            //cvals.Add(theprops[4]);
                            cvals.Add(startLoc.DistanceTo(endLoc).ToString());

                            dlm.Relate("P3dLineGroupPartRelationship", "LineGroup", lgrowId, "Part", crowId);
                            dlm.SetProperties(crowId, cnames, cvals);


                        }
                        else if (theprops[0].Equals("Elbow"))
                        {

                            InlineAsset entelbow = null;

                            using (Transaction tr = db.TransactionManager.StartTransaction())
                            {
                                entelbow = tr.GetObject(id, OpenMode.ForWrite) as InlineAsset;

                                if (!entelbow.ZAxis.Equals(zVector))
                                {
                                    Autodesk.ProcessPower.PnP3dObjects.PortCollection theports = entelbow.GetPorts(PortType.Both);
                                    Vector3d port1 = theports[0].Direction;
                                    Vector3d port2 = theports[1].Direction;
                                    Point3d cog_old = entelbow.CenterOfGravity;
                                    Vector3d theXaxis = port1;
                                    Vector3d theXaxis_alternative = port2;
                                    entelbow.SetOrientation(theXaxis, zVector);
                                    if (!entelbow.CenterOfGravity.Equals(cog_old))
                                        entelbow.SetOrientation(theXaxis_alternative, zVector);
                                }

                                tr.Commit();
                            }

                        }

                    }

                }

            }
            catch (System.Exception ex)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                ed.WriteMessage(trace.ToString());
                ed.WriteMessage("Line: " + trace.GetFrame(0).GetFileLineNumber());
                ed.WriteMessage("message: " + ex.Message);

            }
            finally
            {

            }
        }


        public static string setSpecData(ref Editor ed, Point3d startLoc, Point3d endLoc, string shortdesc, NominalDiameter nd, string H, string pathSpec, ref Database db)
        {


            double dist = startLoc.DistanceTo(endLoc);



            PnPRow specrow = null;

            Autodesk.ProcessPower.PartsRepository.Specification.PipePartSpecification pps = Autodesk.ProcessPower.PartsRepository.Specification.PipePartSpecification.OpenSpecification(pathSpec);

            PnPTable table = pps.Database.Tables["EngineeringItems"];

            if (table != null)
            {
                String query = "\"NominalDiameter\"=" + nd.Value;

                query += " and \"ShortDescription\"='" + shortdesc + "'";

                PnPRow[] r = table.Select(query);

                if (r.Length > 0)
                {
                    specrow = r[0];
                }

                string currentgeometry = specrow["ContentGeometryParamDefinition"].ToString();
                string[] cgeoArr = currentgeometry.Split(new char[] { ',' });
                string newgeometry = "";

                foreach (string cgeo in cgeoArr)
                {
                    if (!newgeometry.Equals(""))
                        newgeometry += ",";
                    if (cgeo.IndexOf("L=") != -1)
                        newgeometry += "L=" + dist;
                    else
                        newgeometry += cgeo;
                    if (cgeo.IndexOf("H=") != -1)
                        H = cgeo.Split(new char[] { '=' })[1].Trim();
                }

                specrow.BeginEdit();
                specrow.SetPropertyValue("ContentGeometryParamDefinition", newgeometry);
                specrow.EndEdit();
                table.RefreshCachedRows();
                pps.AcceptChanges();
                pps.Close();

            }
            return H;
        }

        public static ObjectId insertSpecPartTo3d(Point3d startLoc, Point3d endLoc, Vector3d zVector, ref Editor ed, NominalDiameter nd, string shortdesc, string cspec, ref Database db, ref Project currentProject, ref DataLinksManager dlm)
        {
            ContentManager cm = ContentManager.GetContentManager();
            DataLinksManager3d dlm3d = DataLinksManager3d.Get3dManager(dlm);
            PipingObjectAdder pipeObjAdder = new PipingObjectAdder(dlm3d, db);

            SpecPart specpart_Flange1 = FetchSpecPart("Coupling", shortdesc, cspec, nd, ref ed);
            //SpecPart specpart_Flange1 = FetchSpecPart("Sleeve", shortdesc, cspec, nd, ref ed);

            ObjectId symboloid_Flange1 = ContentManager.GetContentManager().GetSymbol(specpart_Flange1, db);


            // Create the FLANGE 1 part
            //
            PipeInlineAsset pipeInlineAsset_Flange1 = CreateInlineAsset();

            // Set  the FLANGE 1 part's orientation
            //
            pipeInlineAsset_Flange1.Position = startLoc;
            pipeInlineAsset_Flange1.SetOrientation(startLoc.GetVectorTo(endLoc), zVector);
            //pipeInlineAsset_Flange1.SetOrientation(startLoc.GetVectorTo(endLoc), startLoc.GetVectorTo(endLoc).GetPerpendicularVector());

            pipeInlineAsset_Flange1.SymbolId = symboloid_Flange1;

            // Add the FLANGE 1 to the database
            //

            ObjectId oid_Flange1 = AddObjectToDatabase(specpart_Flange1, pipeInlineAsset_Flange1, String.Empty, null, ref db, ref ed, ref currentProject, ref dlm, ref dlm3d, ref pipeObjAdder);

            return oid_Flange1;
        }

    }
}
