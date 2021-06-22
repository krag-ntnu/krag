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
  private void RunScript(string ID, string Master, string Slave, double Tol, Brep Face, ref object Part, ref object INP)
  {
        //Create output tree
    DataTree <System.Object> tree = new DataTree<System.Object>();
    List< string > INPLst = new List<string>();
    if (string.IsNullOrEmpty(ID)) //if ID is not assigned
    {
      ID = "TIE1"; //Make default ID
    }
    //Add surface and ID to be used in Part component
    tree.Add(Face, new GH_Path(0));
    tree.Add(ID, new GH_Path(1));
    //Add tolerance for pointcompare
    if(Tol == 0)
    {
      tree.Add(0.001, new GH_Path(2));
    }
    else
    {
      tree.Add(Tol, new GH_Path(2));
    }
    //Add conditions to be exported in INP
    INPLst.Add(string.Format("** Constraint: {0}", ID));
    INPLst.Add(string.Format("*Tie, name={0}, adjust=yes, type=SURFACE TO SURFACE", ID));
    INPLst.Add(string.Format("{0}-{1}-TieSrf, {0}-{2}-TieSrf", ID, Slave, Master));
    ///////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////
    Part = tree;
    INP = INPLst;
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
        string ID = default(string);
    if (inputs[0] != null)
    {
      ID = (string)(inputs[0]);
    }

    string Master = default(string);
    if (inputs[1] != null)
    {
      Master = (string)(inputs[1]);
    }

    string Slave = default(string);
    if (inputs[2] != null)
    {
      Slave = (string)(inputs[2]);
    }

    double Tol = default(double);
    if (inputs[3] != null)
    {
      Tol = (double)(inputs[3]);
    }

    Brep Face = default(Brep);
    if (inputs[4] != null)
    {
      Face = (Brep)(inputs[4]);
    }



    //3. Declare output parameters
      object Part = null;
  object INP = null;


    //4. Invoke RunScript
    RunScript(ID, Master, Slave, Tol, Face, ref Part, ref INP);
      
    try
    {
      //5. Assign output parameters to component...
            if (Part != null)
      {
        if (GH_Format.TreatAsCollection(Part))
        {
          IEnumerable __enum_Part = (IEnumerable)(Part);
          DA.SetDataList(0, __enum_Part);
        }
        else
        {
          if (Part is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(Part));
          }
          else
          {
            //assign direct
            DA.SetData(0, Part);
          }
        }
      }
      else
      {
        DA.SetData(0, null);
      }
      if (INP != null)
      {
        if (GH_Format.TreatAsCollection(INP))
        {
          IEnumerable __enum_INP = (IEnumerable)(INP);
          DA.SetDataList(1, __enum_INP);
        }
        else
        {
          if (INP is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(1, (Grasshopper.Kernel.Data.IGH_DataTree)(INP));
          }
          else
          {
            //assign direct
            DA.SetData(1, INP);
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
