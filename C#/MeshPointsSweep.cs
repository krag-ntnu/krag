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
  private void RunScript(List<Point3d> PointsSource, List<Point3d> PointsTarget, int Divisions, int U, int V, ref object Mesh, ref object Srf, ref object SourceSrf, ref object TargetSrf, ref object Orientation)
  {
        //riktig oppdatert0306
    List<double> length = new List<double>();
    List <Curve> edgecrv = new List<Curve>();
    DataTree<Curve> edgecrvtree = new DataTree<Curve>();
    DataTree <Point3d> vrttree = new DataTree<Point3d>();
    DataTree <Point3d> vrtsectiontree = new DataTree<Point3d>();
    List <NurbsSurface> s = new List<NurbsSurface>();
    DataTree <System.Object> meshtree = new DataTree<System.Object>();


    //Default values if not specified, added 27.04
    if(Divisions == 0){Divisions = 1;};
    if(U == 0){U = 1;};
    if(V == 0){V = 1;};

    //Make edgecurves, get their length and make new corner points
    for ( int i = 0; i < 4; i++)
    {
      edgecrv.Add(new Line(PointsSource[i], PointsTarget[i]).ToNurbsCurve()); //Add for only one item
      length.Add(edgecrv[i].GetLength());
      edgecrvtree.Add(new Line(PointsSource[i], PointsTarget[i]).ToNurbsCurve(), new GH_Path(i));
      vrttree.AddRange(edgecrvtree.Branch(i)[0].DivideEquidistant(length[i] / Divisions), new GH_Path(i));

      int listitem = 0;
      for (int j = 0; j < (Divisions + 1); j++)
      {
        vrtsectiontree.Add(vrttree.Branch(i)[listitem], new GH_Path(j));
        listitem++;
      }
    }
    //make surfaces
    //get beginning and end surface if needed for tie/boundary conditions or etc.
    Brep ss = Brep.CreateFromCornerPoints(PointsSource[0], PointsSource[1], PointsSource[2], PointsSource[3], 0);
    Brep ts = Brep.CreateFromCornerPoints(PointsTarget[0], PointsTarget[1], PointsTarget[2], PointsTarget[3], 0);

    //get cross-sectional surfaces
    for (int j = 0; j < (Divisions + 1); j++)
    {
      s.Add(NurbsSurface.CreateFromCorners(vrtsectiontree.Branch(j)[0], vrtsectiontree.Branch(j)[1], vrtsectiontree.Branch(j)[2], vrtsectiontree.Branch(j)[3], 0));
    } // (branch number) [list number]


    DataTree<Brep> btree = new DataTree<Brep>();
    DataTree <Curve> loftcrv = new DataTree<Curve>(); //loft to get solid
    DataTree <Point3d> nodetree = new DataTree<Point3d>();


    for (int a = 0; a < (Divisions + 1); a++)
    {
      for (int i = 0; i < (V + 1); i++)
      {
        for (int j = 0; j < (U + 1); j++)
        {
          Point3d p = new Point3d(s[a].PointAt(s[a].Domain(0).Length / V * i, s[a].Domain(1).Length / U * j)); // v-value first, then u value, because domain(0) is v
          meshtree.Add(p, new GH_Path(0, 0)); //added 1704
          //meshtree.Add(p, new GH_Path(1, a)); //added1704 to get nodetree2 in same list
          nodetree.Add(p, new GH_Path(a, i));
        }
      }
    }

    for (int a = 0; a < (Divisions + 1); a++)
    {
      for (int i = 0; i < V; i++)
      {
        for (int j = 0; j < U; j++)
        {
          Brep b = Brep.CreateFromCornerPoints(nodetree.Branch(a, i)[j], nodetree.Branch(a, i + 1)[j], nodetree.Branch(a, i + 1)[j + 1], nodetree.Branch(a, i)[j + 1], 0);
          btree.Add(b, new GH_Path(a, i));
          //meshtree.Add(b, new GH_Path(2, a));
          meshtree.Add(b, new GH_Path(1, 0));
          loftcrv.AddRange(btree.Branch(a, i)[j].DuplicateEdgeCurves(), new GH_Path(a));
        }
      }
    }


    DataTree <Brep> srf = new DataTree<Brep>();
    for ( int i = 0; i < loftcrv.Branch(0).Count; i++)
    {
      for (int j = 0; j < Divisions; j++)
      {
        srf.AddRange(Brep.CreateFromLoft(new List<Curve>{loftcrv.Branch(0 + j)[i], loftcrv.Branch(1 + j)[i]}, Point3d.Unset, Point3d.Unset, LoftType.Normal, false), new GH_Path(i));
      }
    } // Surfaces in the middle are doubled


    //orientation

    Point3d p1 = new Point3d((PointsSource[0].X + PointsSource[1].X + PointsSource[2].X + PointsSource[3].X) / 4, (PointsSource[0].Y + PointsSource[1].Y + PointsSource[2].Y + PointsSource[3].Y) / 4, (PointsSource[0].Z + PointsSource[1].Z + PointsSource[2].Z + PointsSource[3].Z) / 4);
    Point3d p2 = new Point3d((PointsTarget[0].X + PointsTarget[1].X + PointsTarget[2].X + PointsTarget[3].X) / 4, (PointsTarget[0].Y + PointsTarget[1].Y + PointsTarget[2].Y + PointsTarget[3].Y) / 4, (PointsTarget[0].Z + PointsTarget[1].Z + PointsTarget[2].Z + PointsTarget[3].Z) / 4);

    //orientation

    List <string> orilist = new List<string>();

    Vector3d vec1 = p2 - p1; //longitudinal direction of timber cross-section central points = Direction 1 (x-axis). Created between the average points (center points p1 and p2) of the source and target points.
    vec1.Unitize();
    Plane nplan = new Plane(p1, vec1); //normal plane
    Vector3d vec2 = nplan.XAxis; // Second axis (unitized)

    double cx = p1.X; double cy = p1.Y; double cz = p1.Z; //center point x,y,z

    //coordinates of the first principal axis(x-axis)
    double x1 = vec1.X + cx; double y1 = vec1.Y + cy; double z1 = vec1.Z + cz;

    //coordinates of the second principal axis(z-axis in Rhino, y-axis in Abaqus)
    double x2 = vec2.X + cx; double y2 = vec2.Y + cy; double z2 = vec2.Z + cz;

    //Add orientation (for orthotropic material)
    orilist.Add("Ori-1");
    orilist.Add("*Orientation, name=Ori-1");
    orilist.Add(string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}", x1, z1, y1, x2, z2, y2, cx, cz, cy)); //changed y z due to abaqus/rhino different axis systems
    orilist.Add("1, 0"); // 1,0. 1 means primary axis (x), 0 means no rotation


    //add divisions, u and v to be used further to end of mesh tree
    GH_Path pth = new GH_Path(2, 0);
    meshtree.Add(Divisions, pth); //number of divisions
    //meshtree.Add((U + 1) * (V + 1), pth); // number of nodes per branch
    meshtree.Add(U * V, pth); //number of surfaces per branch


    Srf = srf;
    Mesh = meshtree;
    SourceSrf = ss;
    TargetSrf = ts;
    Orientation = orilist;





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
        List<Point3d> PointsSource = null;
    if (inputs[0] != null)
    {
      PointsSource = GH_DirtyCaster.CastToList<Point3d>(inputs[0]);
    }
    List<Point3d> PointsTarget = null;
    if (inputs[1] != null)
    {
      PointsTarget = GH_DirtyCaster.CastToList<Point3d>(inputs[1]);
    }
    int Divisions = default(int);
    if (inputs[2] != null)
    {
      Divisions = (int)(inputs[2]);
    }

    int U = default(int);
    if (inputs[3] != null)
    {
      U = (int)(inputs[3]);
    }

    int V = default(int);
    if (inputs[4] != null)
    {
      V = (int)(inputs[4]);
    }



    //3. Declare output parameters
      object Mesh = null;
  object Srf = null;
  object SourceSrf = null;
  object TargetSrf = null;
  object Orientation = null;


    //4. Invoke RunScript
    RunScript(PointsSource, PointsTarget, Divisions, U, V, ref Mesh, ref Srf, ref SourceSrf, ref TargetSrf, ref Orientation);
      
    try
    {
      //5. Assign output parameters to component...
            if (Mesh != null)
      {
        if (GH_Format.TreatAsCollection(Mesh))
        {
          IEnumerable __enum_Mesh = (IEnumerable)(Mesh);
          DA.SetDataList(0, __enum_Mesh);
        }
        else
        {
          if (Mesh is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(Mesh));
          }
          else
          {
            //assign direct
            DA.SetData(0, Mesh);
          }
        }
      }
      else
      {
        DA.SetData(0, null);
      }
      if (Srf != null)
      {
        if (GH_Format.TreatAsCollection(Srf))
        {
          IEnumerable __enum_Srf = (IEnumerable)(Srf);
          DA.SetDataList(1, __enum_Srf);
        }
        else
        {
          if (Srf is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(1, (Grasshopper.Kernel.Data.IGH_DataTree)(Srf));
          }
          else
          {
            //assign direct
            DA.SetData(1, Srf);
          }
        }
      }
      else
      {
        DA.SetData(1, null);
      }
      if (SourceSrf != null)
      {
        if (GH_Format.TreatAsCollection(SourceSrf))
        {
          IEnumerable __enum_SourceSrf = (IEnumerable)(SourceSrf);
          DA.SetDataList(2, __enum_SourceSrf);
        }
        else
        {
          if (SourceSrf is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(2, (Grasshopper.Kernel.Data.IGH_DataTree)(SourceSrf));
          }
          else
          {
            //assign direct
            DA.SetData(2, SourceSrf);
          }
        }
      }
      else
      {
        DA.SetData(2, null);
      }
      if (TargetSrf != null)
      {
        if (GH_Format.TreatAsCollection(TargetSrf))
        {
          IEnumerable __enum_TargetSrf = (IEnumerable)(TargetSrf);
          DA.SetDataList(3, __enum_TargetSrf);
        }
        else
        {
          if (TargetSrf is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(3, (Grasshopper.Kernel.Data.IGH_DataTree)(TargetSrf));
          }
          else
          {
            //assign direct
            DA.SetData(3, TargetSrf);
          }
        }
      }
      else
      {
        DA.SetData(3, null);
      }
      if (Orientation != null)
      {
        if (GH_Format.TreatAsCollection(Orientation))
        {
          IEnumerable __enum_Orientation = (IEnumerable)(Orientation);
          DA.SetDataList(4, __enum_Orientation);
        }
        else
        {
          if (Orientation is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(4, (Grasshopper.Kernel.Data.IGH_DataTree)(Orientation));
          }
          else
          {
            //assign direct
            DA.SetData(4, Orientation);
          }
        }
      }
      else
      {
        DA.SetData(4, null);
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
