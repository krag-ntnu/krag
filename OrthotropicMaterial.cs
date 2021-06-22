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
  private void RunScript(string Material, int E1, int E2, int E3, double Nu12, double Nu13, double Nu23, int G12, int G13, int G23, ref object Properties)
  {
    
    List< System.Object > B = new List<System.Object>();
    ///////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////
    if (string.IsNullOrEmpty(Material)){Material = "Timber";}
    B.Add(Material);
    B.Add("Elastic, type=ENGINEERING CONSTANTS");
    //Elastic moduli
    if(E1 == 0){E1 = 10000;}
    if(E2 == 0){E2 = 800;}
    if(E3 == 0){E3 = 400;}

    //Poisson's ratios
    if(Nu12 == 0){Nu12 = 0.5;}
    if(Nu13 == 0){Nu13 = 0.6;}
    if(Nu23 == 0){Nu23 = 0.6;}

    //Shear moduli
    if(G12 == 0){G12 = 600;}
    if(G13 == 0){G13 = 600;}
    if(G23 == 0){G23 = 30;}


    B.Add(string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", E1, E2, E3, Nu12, Nu13, Nu23, G12, G13));
    B.Add(string.Format("{0}", G23)); //Apparently the ABAQUS demands that G23 is on the second line.

    Properties = B;


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
        string Material = default(string);
    if (inputs[0] != null)
    {
      Material = (string)(inputs[0]);
    }

    int E1 = default(int);
    if (inputs[1] != null)
    {
      E1 = (int)(inputs[1]);
    }

    int E2 = default(int);
    if (inputs[2] != null)
    {
      E2 = (int)(inputs[2]);
    }

    int E3 = default(int);
    if (inputs[3] != null)
    {
      E3 = (int)(inputs[3]);
    }

    double Nu12 = default(double);
    if (inputs[4] != null)
    {
      Nu12 = (double)(inputs[4]);
    }

    double Nu13 = default(double);
    if (inputs[5] != null)
    {
      Nu13 = (double)(inputs[5]);
    }

    double Nu23 = default(double);
    if (inputs[6] != null)
    {
      Nu23 = (double)(inputs[6]);
    }

    int G12 = default(int);
    if (inputs[7] != null)
    {
      G12 = (int)(inputs[7]);
    }

    int G13 = default(int);
    if (inputs[8] != null)
    {
      G13 = (int)(inputs[8]);
    }

    int G23 = default(int);
    if (inputs[9] != null)
    {
      G23 = (int)(inputs[9]);
    }



    //3. Declare output parameters
      object Properties = null;


    //4. Invoke RunScript
    RunScript(Material, E1, E2, E3, Nu12, Nu13, Nu23, G12, G13, G23, ref Properties);
      
    try
    {
      //5. Assign output parameters to component...
            if (Properties != null)
      {
        if (GH_Format.TreatAsCollection(Properties))
        {
          IEnumerable __enum_Properties = (IEnumerable)(Properties);
          DA.SetDataList(0, __enum_Properties);
        }
        else
        {
          if (Properties is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(Properties));
          }
          else
          {
            //assign direct
            DA.SetData(0, Properties);
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