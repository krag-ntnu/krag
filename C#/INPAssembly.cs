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
  private void RunScript(List<string> ID, List<string> Job, List<string> Model, List<string> Part, List<string> Instance, List<string> Sets, List<string> BC, List<string> Material, List<string> Tie, List<string> Steps, ref object INP)
  {
        //Create output;
    List < string > INPLst = new List<string>{ "*Heading" };
    //Create variables;
    string job, model;
    //Include the header, job- and model name;Model
    if (ID.Count == 0)
    {
      INPLst.Add("GH_File");
    }
    else
    {
      INPLst.Add(ID[0]);
    }
    if (Job.Count == 0)
    {
      job = "GH_Job";
    }
    else
    {
      job = Job[0];
    }
    if (Model.Count == 0)
    {
      model = "GH_Model";
    }
    else
    {
      model = Model[0];
    }
    INPLst.Add(string.Format("** Job name: {0} Model name: {1}", job, model));
    ///////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////
    //Setting the data file printing options;
    INPLst.Add("*Preprint, echo=NO, model=NO, history=NO, contact=NO");
    //Parts;
    INPLst.Add("**");
    INPLst.Add("** PARTS");
    INPLst.Add("**");
    INPLst.AddRange(Part);
    //Start assemble the model;
    INPLst.Add("**");
    INPLst.Add("** ASSEMBLY");
    INPLst.Add("**");
    INPLst.Add("*Assembly, name=Assembly");
    INPLst.Add("**");
    //Instances;
    INPLst.AddRange(Instance);
    //Sets;
    INPLst.AddRange(Sets);
    //Tie;
    INPLst.AddRange(Tie);
    //End assembly;
    INPLst.Add("*End Assembly");
    //Defining the material properties;
    INPLst.Add("**");
    INPLst.Add("** MATERIALS");
    INPLst.Add("**");

    /*if ( Material.Count > 3)
    {
      INPLst.Add(string.Format("*Material, name={0}", Material[0]));
      INPLst.Add(string.Format("* {0}", Material[1]));
      INPLst.Add(string.Format("{0}, {1}", Material[2], Material[3]));
    }
 */

    for (int i = 0; i < Material.Count; i += 4) //fixed 2904
    {
      INPLst.Add(string.Format("*Material, name={0}", Material[i]));
      INPLst.Add(string.Format("*{0}", Material[i + 1]));
      INPLst.Add(string.Format("{0}", Material[i + 2]));
      INPLst.Add(string.Format("{0}", Material[i + 3]));
    }


    //Defining the boundary conditions;
    INPLst.Add("**");
    INPLst.Add("** BOUNDARY CONDITIONS");
    INPLst.Add("**");
    INPLst.Add("** Name: BC-1 Type: Displacement/Rotation");
    INPLst.Add("*Boundary");
    INPLst.AddRange(BC);
    //End defining BC;
    INPLst.Add("** ----------------------------------------------------------------");
    //Add the steps;
    INPLst.AddRange(Steps);
    ///////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////
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
        List<string> ID = null;
    if (inputs[0] != null)
    {
      ID = GH_DirtyCaster.CastToList<string>(inputs[0]);
    }
    List<string> Job = null;
    if (inputs[1] != null)
    {
      Job = GH_DirtyCaster.CastToList<string>(inputs[1]);
    }
    List<string> Model = null;
    if (inputs[2] != null)
    {
      Model = GH_DirtyCaster.CastToList<string>(inputs[2]);
    }
    List<string> Part = null;
    if (inputs[3] != null)
    {
      Part = GH_DirtyCaster.CastToList<string>(inputs[3]);
    }
    List<string> Instance = null;
    if (inputs[4] != null)
    {
      Instance = GH_DirtyCaster.CastToList<string>(inputs[4]);
    }
    List<string> Sets = null;
    if (inputs[5] != null)
    {
      Sets = GH_DirtyCaster.CastToList<string>(inputs[5]);
    }
    List<string> BC = null;
    if (inputs[6] != null)
    {
      BC = GH_DirtyCaster.CastToList<string>(inputs[6]);
    }
    List<string> Material = null;
    if (inputs[7] != null)
    {
      Material = GH_DirtyCaster.CastToList<string>(inputs[7]);
    }
    List<string> Tie = null;
    if (inputs[8] != null)
    {
      Tie = GH_DirtyCaster.CastToList<string>(inputs[8]);
    }
    List<string> Steps = null;
    if (inputs[9] != null)
    {
      Steps = GH_DirtyCaster.CastToList<string>(inputs[9]);
    }


    //3. Declare output parameters
      object INP = null;


    //4. Invoke RunScript
    RunScript(ID, Job, Model, Part, Instance, Sets, BC, Material, Tie, Steps, ref INP);
      
    try
    {
      //5. Assign output parameters to component...
            if (INP != null)
      {
        if (GH_Format.TreatAsCollection(INP))
        {
          IEnumerable __enum_INP = (IEnumerable)(INP);
          DA.SetDataList(0, __enum_INP);
        }
        else
        {
          if (INP is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(INP));
          }
          else
          {
            //assign direct
            DA.SetData(0, INP);
          }
        }
      }
      else
      {
        DA.SetData(0, null);
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
