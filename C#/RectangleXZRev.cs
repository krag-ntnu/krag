using System;
using System.Collections;
using System.Collections.Generic;

using Rhino;
using Rhino.Geometry;

using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

using System.IO;
using System.Linq;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Runtime.InteropServices;

using Rhino.DocObjects;
using Rhino.Collections;
using GH_IO;
using GH_IO.Serialization;

/// <summary>
/// This class will be instantiated on demand by the Script component.
/// </summary>
public class Script_Instance : GH_ScriptInstance
{
#region Utility functions
  /// <summary>Print a String to the [Out] Parameter of the Script component.</summary>
  /// <param name="text">String to print.</param>
  private void Print(string text) { __out.Add(text); }
  /// <summary>Print a formatted String to the [Out] Parameter of the Script component.</summary>
  /// <param name="format">String format.</param>
  /// <param name="args">Formatting parameters.</param>
  private void Print(string format, params object[] args) { __out.Add(string.Format(format, args)); }
  /// <summary>Print useful information about an object instance to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj) { __out.Add(GH_ScriptComponentUtilities.ReflectType_CS(obj)); }
  /// <summary>Print the signatures of all the overloads of a specific method to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj, string method_name) { __out.Add(GH_ScriptComponentUtilities.ReflectType_CS(obj, method_name)); }
#endregion

#region Members
  /// <summary>Gets the current Rhino document.</summary>
  private RhinoDoc RhinoDocument;
  /// <summary>Gets the Grasshopper document that owns this script.</summary>
  private GH_Document GrasshopperDocument;
  /// <summary>Gets the Grasshopper script component that owns this script.</summary>
  private IGH_Component Component; 
  /// <summary>
  /// Gets the current iteration count. The first call to RunScript() is associated with Iteration==0.
  /// Any subsequent call within the same solution will increment the Iteration count.
  /// </summary>
  private int Iteration;
#endregion

  /// <summary>
  /// This procedure contains the user code. Input parameters are provided as regular arguments, 
  /// Output parameters as ref arguments. You don't have to assign output parameters, 
  /// they will have a default value.
  /// </summary>
  private void RunScript(Point3d Cntr, double W, double H, ref object Rcntr, ref object Pts)
  {
        Point3d firstp = new Point3d(Cntr.X - W / 2, Cntr.Y, Cntr.Z - H / 2);
    Point3d secondp = new Point3d(Cntr.X - W / 2, Cntr.Y, Cntr.Z + H / 2);
    Point3d thirdp = new Point3d(Cntr.X + W / 2, Cntr.Y, Cntr.Z + H / 2);
    Point3d fourthp = new Point3d(Cntr.X + W / 2, Cntr.Y, Cntr.Z - H / 2);
    List<Point3d> pointlist = new List<Point3d>();
    pointlist.Add(fourthp); //reverse points to avoid negative volume
    pointlist.Add(thirdp);
    pointlist.Add(secondp);
    pointlist.Add(firstp);
    Pts = pointlist;
    Rcntr = Cntr;

  }

  // <Custom additional code> 
  
  // </Custom additional code> 

  private List<string> __err = new List<string>(); //Do not modify this list directly.
  private List<string> __out = new List<string>(); //Do not modify this list directly.
  private RhinoDoc doc = RhinoDoc.ActiveDoc;       //Legacy field.
  private IGH_ActiveObject owner;                  //Legacy field.
  private int runCount;                            //Legacy field.
  
  public override void InvokeRunScript(IGH_Component owner, object rhinoDocument, int iteration, List<object> inputs, IGH_DataAccess DA)
  {
    //Prepare for a new run...
    //1. Reset lists
    this.__out.Clear();
    this.__err.Clear();

    this.Component = owner;
    this.Iteration = iteration;
    this.GrasshopperDocument = owner.OnPingDocument();
    this.RhinoDocument = rhinoDocument as Rhino.RhinoDoc;

    this.owner = this.Component;
    this.runCount = this.Iteration;
    this. doc = this.RhinoDocument;

    //2. Assign input parameters
        Point3d Cntr = default(Point3d);
    if (inputs[0] != null)
    {
      Cntr = (Point3d)(inputs[0]);
    }

    double W = default(double);
    if (inputs[1] != null)
    {
      W = (double)(inputs[1]);
    }

    double H = default(double);
    if (inputs[2] != null)
    {
      H = (double)(inputs[2]);
    }



    //3. Declare output parameters
      object Rcntr = null;
  object Pts = null;


    //4. Invoke RunScript
    RunScript(Cntr, W, H, ref Rcntr, ref Pts);
      
    try
    {
      //5. Assign output parameters to component...
            if (Rcntr != null)
      {
        if (GH_Format.TreatAsCollection(Rcntr))
        {
          IEnumerable __enum_Rcntr = (IEnumerable)(Rcntr);
          DA.SetDataList(0, __enum_Rcntr);
        }
        else
        {
          if (Rcntr is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(Rcntr));
          }
          else
          {
            //assign direct
            DA.SetData(0, Rcntr);
          }
        }
      }
      else
      {
        DA.SetData(0, null);
      }
      if (Pts != null)
      {
        if (GH_Format.TreatAsCollection(Pts))
        {
          IEnumerable __enum_Pts = (IEnumerable)(Pts);
          DA.SetDataList(1, __enum_Pts);
        }
        else
        {
          if (Pts is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(1, (Grasshopper.Kernel.Data.IGH_DataTree)(Pts));
          }
          else
          {
            //assign direct
            DA.SetData(1, Pts);
          }
        }
      }
      else
      {
        DA.SetData(1, null);
      }

    }
    catch (Exception ex)
    {
      this.__err.Add(string.Format("Script exception: {0}", ex.Message));
    }
    finally
    {
      //Add errors and messages... 
      if (owner.Params.Output.Count > 0)
      {
        if (owner.Params.Output[0] is Grasshopper.Kernel.Parameters.Param_String)
        {
          List<string> __errors_plus_messages = new List<string>();
          if (this.__err != null) { __errors_plus_messages.AddRange(this.__err); }
          if (this.__out != null) { __errors_plus_messages.AddRange(this.__out); }
          if (__errors_plus_messages.Count > 0) 
            DA.SetDataList(0, __errors_plus_messages);
        }
      }
    }
  }
}
