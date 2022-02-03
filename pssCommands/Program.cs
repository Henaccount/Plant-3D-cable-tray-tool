#region Namespaces

using System;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using System.Collections.Specialized;
using System.Collections.Generic;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.PnP3dObjects;

#endregion

//V0
[assembly: CommandClass(typeof(pssCommands.Program))]


namespace pssCommands
{
    public class Program
    {

        public static bool firstrun = true;


        public static int getOFinfo(String partParams, bool trueIsIndexFalseIsValue)
        {

            int idx = -1;
            IDictionary<string, double> partParamsDict = Helper.ReadInDict(partParams, true);


            foreach (var item in partParamsDict)
            {
                ++idx;
                if (item.Key.Equals("OF"))
                {
                    if (trueIsIndexFalseIsValue)
                        return idx;
                    else
                        return Convert.ToInt32(item.Value);
                }

            }

            return idx;
        }

        public static void setOF(ObjectId id, double length)
        {
            Autodesk.ProcessPower.ACPUtils.ParameterList plist = Autodesk.ProcessPower.PnP3dPipeSupport.SupportHelper.GetSupportParameters(id);
            String partParams = plist.ToString();
            int ofidx = getOFinfo(partParams, true);


            if (ofidx != -1)
            {
                Autodesk.ProcessPower.ACPUtils.ParameterInfo pinfo = new Autodesk.ProcessPower.ACPUtils.ParameterInfo();
                pinfo.Name = "L";
                pinfo.Value = length.ToString();
                plist[ofidx] = pinfo;
                Autodesk.ProcessPower.PnP3dPipeSupport.SupportHelper.UpdateSupportEntity(id, plist);
            }
        }


        [CommandMethod("traysetlength", CommandFlags.UsePickSet)]
        public static void obsOn()
        {

            Helper.Initialize();

            try
            {
                if (!PnPProjectUtils.GetActiveDocumentType().Equals("Piping"))
                {
                    Helper.ed.WriteMessage("\n This tool works only on Piping drawings!");
                    return;
                }

                ObjectId objId = new ObjectId();

                PromptSelectionResult selectionRes = Helper.ed.SelectImplied();
                Point3d pickpoint = new Point3d();

                if (selectionRes.Status == PromptStatus.Error)

                {
                    PromptEntityResult selResult = Helper.ed.GetEntity("Select tray, pick it close to the end where the modification is needed");
                    if (selResult.Status == PromptStatus.OK)
                    {
                        objId = selResult.ObjectId;
                        pickpoint = selResult.PickedPoint;
                    }
                }
                else
                {
                    objId = selectionRes.Value.GetObjectIds()[0];
                }

                PromptPointOptions opt = new PromptPointOptions("\n\nSelect end point or other limit");
                PromptPointResult res;
                do
                {
                    res = Helper.ed.GetPoint(opt);
                }
                while (res.Status == PromptStatus.Error);

                Point3d point1 = res.Value;

                if (point1.Equals(new Point3d())) return;

                /* V0,use two pickpoints:
                 * opt = new PromptPointOptions("Select second point");
                 do
                 {
                     res = Helper.ed.GetPoint(opt);
                 }
                 while (res.Status == PromptStatus.Error);

                 Point3d point2 = res.Value;*/


                if (firstrun)
                {
                    Helper.ed.WriteMessage("\n AUTODESK PROVIDES THIS PROGRAM \"AS IS\" AND WITH ALL FAULTS.");
                    Helper.ed.WriteMessage("\n AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF");
                    Helper.ed.WriteMessage("\n MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.");
                    Helper.ed.WriteMessage("\n DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE");
                    Helper.ed.WriteMessage("\n UNINTERRUPTED OR ERROR FREE.");
                    firstrun = false;
                }

                using (Transaction tr = Helper.db.TransactionManager.StartTransaction())
                {

                    try
                    {
                        InlineAsset tray = tr.GetObject(objId, OpenMode.ForWrite) as InlineAsset;
                        Part asset = tr.GetObject(objId, OpenMode.ForRead) as Part;
                        PortCollection portCol = asset.GetPorts(PortType.All);
                        double length = 0.0;





                        Vector3d S1toEndpoint = tray.Position.GetVectorTo(point1);
                        Vector3d S1toS2 = tray.Position.GetVectorTo(portCol[1].Position);

                        //V2,legacy: if direction of vector s1,endpoint not equals direction s1,s2 (dotproduct < 0) then set tray position to endpoint 
                        //double dotprod = S1toEndpoint.DotProduct(S1toS2);


                        if (pickpoint.DistanceTo(portCol[1].Position) > pickpoint.DistanceTo(tray.Position))
                        {
                            double oldlength = portCol[1].Position.DistanceTo(tray.Position);
                            //V2:without snapping use pickpoint: length = portCol[1].Position.DistanceTo(point1);
                            //use snaps outside of tray axis: length = dotproduct of (vector of the cable tray p1 -> p2) and (vector basepoint of tray to pickpoint)
                            length = portCol[1].Position.GetVectorTo(tray.Position).GetNormal().DotProduct(portCol[1].Position.GetVectorTo(point1));
                            if (length <= 0)
                            {
                                Helper.ed.WriteMessage("\nerror: length cannot be zero or negative");
                                return;
                            }
                            //normalized vector p2p1 * lengthdiff
                            Vector3d moveVector = portCol[1].Position.GetVectorTo(tray.Position).GetNormal().MultiplyBy(length - oldlength);
                            tray.Position = tray.Position.TransformBy(Matrix3d.Displacement(moveVector));
                        }
                        else
                        {
                            //V2:without snapping use pickpoint: length = tray.Position.DistanceTo(point1);
                            length = tray.Position.GetVectorTo(portCol[1].Position).GetNormal().DotProduct(tray.Position.GetVectorTo(point1));
                            if (length <= 0)
                            {
                                Helper.ed.WriteMessage("\nerror: length cannot be zero or negative");
                                return;
                            }
                        }
                        setOF(objId, length);

                    }
                    catch (System.Exception e)
                    {
                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                        Helper.ed.WriteMessage(trace.ToString());
                        Helper.ed.WriteMessage("\nLine: " + trace.GetFrame(0).GetFileLineNumber());
                        Helper.ed.WriteMessage("\nitem error: " + e.Message);
                    }


                    tr.Commit();


                    Helper.ed.WriteMessage("\nScript finished");
                }
            }
            catch (System.Exception e)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.ed.WriteMessage(trace.ToString());
                Helper.ed.WriteMessage("\nLine: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.ed.WriteMessage("\nscript error: " + e.Message);
            }
            finally
            {
                Helper.Terminate();
            }
        }


    }




}

