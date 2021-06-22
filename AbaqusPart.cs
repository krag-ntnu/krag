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
  private void RunScript(string ID, DataTree<System.Object> Mesh, List<string> ElmntType, List<System.Object> Material, List<string> Orientation, DataTree<System.Object> Support, DataTree<System.Object> Load, DataTree<System.Object> Tie, bool Run, ref object Geometry, ref object Part, ref object Instance, ref object Sets, ref object BC, ref object Steps)
  {
        //oppdatert 0306
    //Declare variables;
    Point3d Crd;
    string Nr, X, Y, Z;

    if (string.IsNullOrEmpty(ID)) //if ID is not assigned
    {
      ID = "Part1"; //Make default ID
    }

    //Create output tree;
    DataTree <Curve> Crvs = new DataTree<Curve>();
    DataTree <Point3d> Elements = new DataTree<Point3d>();
    DataTree <System.Object> GeoTree = new DataTree<System.Object>();
    DataTree <string> Inf = new DataTree<string>();
    DataTree <int> Tick = new DataTree<int>();
    DataTree <int> Tock = new DataTree<int>();

    //Create Lists;
    List< Point3d > Vertices = new List<Point3d>();
    List< Point3d > Nds = new List<Point3d>();
    List< Brep > Srf = new List<Brep>();
    List< Brep > ObjList = new List<Brep>();
    List< string > PrtLst = new List<string>();
    List< string > SetLst = new List<string>();
    List< string > InsLst = new List<string>();
    List< string > BCLst = new List<string>();
    List< string > StpLst = new List<string>();
    List< string > ElNr = new List<string>();
    List<Brep> SupLst = new List<Brep>();
    List<double> TolSupLst = new List<double>();
    List<Brep> DisLst = new List<Brep>();
    List<double> TolDisLst = new List<double>();
    List<Brep> LdLst = new List<Brep>();
    List<double> TolLdLst = new List<double>();
    List<Brep> TieLst = new List<Brep>();
    List<double> TolTieLst = new List<double>();

    ////////////////////////////////////////////////////////////////////////////////////
    //Sorts number of meshes into correct lists so that it is possible to assign different values of v,u and div to each mesh and merge meshes into same part
    List<Point3d> nodelist = new List<Point3d>();
    DataTree<Point3d> nodetree = new DataTree<Point3d>();
    DataTree<Brep> breps = new DataTree<Brep>();

    for (int i = 0; i < Mesh.Branch(0, 0).Count(); i++)
    {
      Point3d n = (Point3d) Mesh.Branch(0, 0)[i];
      nodelist.Add(n);
    }

    int div, numbr;
    int startlocb = 0; //startindex of brep
    int count = 0; //count to get correct assignment of div,numnb,numbr;
    for (int j = 0; j < (Mesh.Branch(2, 0).Count() / 2) ; j++) //number of meshes
    {
      div = (int) Mesh.Branch(2, 0)[count]; //divisions j mesh
      numbr = (int) Mesh.Branch(2, 0)[count + 1]; //number of surfaces in each division
      for (int k = 0; k < (div + 1); k++)
      {
        for (int p = startlocb; p < (startlocb + numbr); p++)
        {
          Brep bb = (Brep) Mesh.Branch(1, 0)[p];
          breps.Add(bb, new GH_Path(j, k));
        }
        for (int q = 0; q < numbr; q++)
        {
          Elements.AddRange(breps.Branch(j, k)[q].DuplicateVertices(), new GH_Path(j, k, q));
        }
        startlocb = startlocb + numbr; //make next location
      }
      count = count + 2; //to get next mesh's division and number of surfaces.
    }



    ///////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////////////////////

    if (Run)
    {
      //Discretize the data of the Part;
      PrtLst.Add(string.Format("*Part, name={0}", ID));
      //Nodes of the Part;
      PrtLst.Add("*Node");

      Nds.AddRange(Point3d.CullDuplicates(nodelist, doc.ModelAbsoluteTolerance)); //remove duplicate node points
      for ( int i = 0; i < Nds.Count; i++)
      {
        Nr = (i + 1).ToString();
        Crd = Nds[i];
        X = Crd.X.ToString();
        Y = Crd.Y.ToString();
        Z = Crd.Z.ToString();
        PrtLst.Add(Nr + ", " + X + ", " + Z + ", " + Y);
      }
      //Sort element Nodes and identify faces which is supported or loaded;
      PrtLst.Add(string.Format("*Element, type={0}", ElmntType[0]));

      //If Load, import input for load;
      if ( Load.PathExists(0) )
      {
        double TolLd;
        Brep LdSrf = new Brep();
        for (int i = 0; i < Load.Branch(0).Count(); i++)
        {
          GH_Convert.ToBrep(Load.Branch(0)[i], ref LdSrf, Grasshopper.Kernel.GH_Conversion.Both);
          GH_Convert.ToDouble(Load.Branch(3)[i], out TolLd, GH_Conversion.Both);
          TolLdLst.Add(TolLd);
          LdLst.Add(LdSrf);
        }
      }
      //If Tie, import input for tie;
      if ( Tie.PathExists(0) )
      {
        double TolTie;
        Brep TieSrf = new Brep();
        for (int i = 0; i < Tie.Branch(0).Count(); i++)
        {
          GH_Convert.ToBrep(Tie.Branch(0)[i], ref TieSrf, Grasshopper.Kernel.GH_Conversion.Both);
          GH_Convert.ToDouble(Tie.Branch(2)[i], out TolTie, GH_Conversion.Both);
          TolTieLst.Add(TolTie);
          TieLst.Add(TieSrf);
        }
      }

      //Make list of elements, identify load and tie
      int count2 = 0; //To get the correct assignment of divisions, number of breps for each mesh
      int count3 = 1; //Begin at element 1 in Abaqus

      for (int f = 0; f < (Mesh.Branch(2, 0).Count() / 2) ; f++) //number of meshes
      {
        int div2 = (int) Mesh.Branch(2, 0)[count2]; //divisions of f mesh
        int numbr2 = (int) Mesh.Branch(2, 0)[count2 + 1]; //numbers of breps in each division

        for (int g = 0; g < div2; g++)
        {
          for (int h = 0; h < numbr2; h++)
          {
            ElNr.Add(count3.ToString());
            for ( int k = 0; k < Elements.Branch(f, g, h).Count; k++)
            {
              for ( int l = 0; l < Nds.Count; l++)
              {
                if ( PtCmp(Elements.Branch(f, g, h)[k], Nds[l]))
                {
                  ElNr.Add((l + 1).ToString());
                }
              } //4 first corners
            }
            for ( int k = 0; k < Elements.Branch(f, g, h).Count; k++)
            {
              for ( int l = 0; l < Nds.Count; l++)
              {
                if ( PtCmp(Elements.Branch(f, g + 1, h)[k], Nds[l]))
                {
                  ElNr.Add((l + 1).ToString());
                }
              } //4 second corners

              //identify if srf is loaded
              if ( Load.PathExists(0) )
              {
                for (int m = 0; m < Load.Branch(0).Count; m++)
                {
                  Tick.EnsurePath(m);
                  if ( IdenSet(LdLst[m], Elements.Branch(f, g, h)[k], TolLdLst[m]))
                  {
                    Tick.Add(k + 1, new GH_Path(m));
                    if (Tick.Branch(m).Count == 4)
                    {
                      Inf.Add((count3).ToString() + ",", new GH_Path(1, m, 2));
                      Inf.Add(SrfInd(Tick.Branch(m)).ToString(), new GH_Path(1, m, 3));
                    }
                  }
                  if ( IdenSet(LdLst[m], Elements.Branch(f, g + 1, h)[k], TolLdLst[m]))
                  {
                    Tick.Add(k + 5, new GH_Path(m));
                    if (Tick.Branch(m).Count == 4)
                    {
                      Inf.Add((count3).ToString() + ",", new GH_Path(1, m, 2));
                      Inf.Add(SrfInd(Tick.Branch(m)).ToString(), new GH_Path(1, m, 3));
                    }
                  }
                }
              } //end load Path

              //identify if nodes are tied
              if ( Tie.PathExists(0) )
              {
                for (int m = 0; m < Tie.Branch(0).Count; m++)
                {
                  Tock.EnsurePath(m);
                  if ( IdenSet(TieLst[m], Elements.Branch(f, g, h)[k], TolTieLst[m]))
                  {
                    Tock.Add(k + 1, new GH_Path(m));
                    if (Tock.Branch(m).Count == 4)
                    {
                      Inf.Add((count3).ToString() + ",", new GH_Path(1, m, 4));
                      Inf.Add(SrfInd(Tock.Branch(m)).ToString(), new GH_Path(1, m, 5));
                    }
                  }
                  if ( IdenSet(TieLst[m], Elements.Branch(f, g + 1, h)[k], TolTieLst[m]))
                  {
                    Tock.Add(k + 5, new GH_Path(m));
                    if (Tock.Branch(m).Count == 4)
                    {
                      Inf.Add((count3).ToString() + ",", new GH_Path(1, m, 4));
                      Inf.Add(SrfInd(Tock.Branch(m)).ToString(), new GH_Path(1, m, 5));
                    }
                  }
                }
              } //end tie
            }

            PrtLst.Add(String.Join(",  ", ElNr.ToArray()));
            count3 = count3 + 1;
            ElNr.Clear();
            Tick.ClearData();
            Tock.ClearData();
          }
        }
        count2 = count2 + 2;
      }


      //Identify nodes on supported face;
      if ( Support.PathExists(0) )
      {
        double TolSup;
        Brep SupSrf = new Brep();
        for (int i = 0; i < Support.Branch(0).Count(); i++)
        {
          GH_Convert.ToBrep(Support.Branch(0)[i], ref SupSrf, Grasshopper.Kernel.GH_Conversion.Both);
          GH_Convert.ToDouble(Support.Branch(3)[i], out TolSup, GH_Conversion.Both);
          TolSupLst.Add(TolSup);
          SupLst.Add(SupSrf);
        }
        for (int i = 0; i < Nds.Count; i++)
        {
          for (int j = 0; j < Support.Branch(0).Count; j++)
          {
            if ( IdenSet(SupLst[j], Nds[i], TolSupLst[j]) )
            {
              Inf.Add((i + 1).ToString() + ",", new GH_Path(0, j, 1));
              GeoTree.Add(Nds[i].X.ToString(), new GH_Path(2, j, 0));
              GeoTree.Add(Nds[i].Y.ToString(), new GH_Path(2, j, 1));
              GeoTree.Add(Nds[i].Z.ToString(), new GH_Path(2, j, 2));
            }
          }
        }
      }

      ///////////////////////////////////////////////////////////////////////////////////////
      ///////////////////////////////////////////////////////////////////////////////////////
      //Create sets of nodes;
      PrtLst.Add("*Nset, nset=Set-1, generate");
      PrtLst.Add(string.Format("   1,    {0},    1", Nds.Count));
      //Create sets of elements;
      PrtLst.Add("*Elset, elset=Set-1, generate");
      PrtLst.Add(string.Format("   1,    {0},    1", count3 - 1));

      //Add orientation (for orthotropic material) 30.04

      if ( (Orientation != null) && (!Orientation.Any()) ) //if the list is empty
      {
        PrtLst.Add(string.Format("**Section: {0}_section", ID));
        PrtLst.Add(string.Format("*Solid Section, elset=Set-1, material={0}", Material[0]));
      }
      else
      {
        PrtLst.Add(Orientation[1]);
        PrtLst.Add(Orientation[2]);
        PrtLst.Add(Orientation[3]);
        PrtLst.Add(string.Format("**Section: {0}_section", ID));
        PrtLst.Add(string.Format("*Solid Section, elset=Set-1, orientation=Ori-1, material={0}", Material[0]));
      }

      PrtLst.Add(",");
      //Finish defining Part;
      PrtLst.Add("*End Part");
      PrtLst.Add("**");
      //Create instance of the part;
      InsLst.Add(string.Format("*Instance, name={0}, part={0}", ID));
      InsLst.Add("*End Instance");
      InsLst.Add("**");
      //BC; Add nodes on supported Face;
      if ( Inf.PathExists(0, 0, 1) )
      {
        for (int i = 0; i < Support.Branch(0).Count; i++)
        {
          SetLst.Add(string.Format("*Nset, nset={0}, instance={1}", Support.Branch(2)[i], ID));
          for (int j = 0; j < Inf.Branch(0, i, 1).Count; j++)
          {
            SetLst.Add(Inf.Branch(0, i, 1)[j]);
          }
        }
      }
      //Add elements on tied face;
      if ( Inf.PathExists(1, 0, 4) )
      {
        for (int i = 0; i < Tie.Branch(0).Count; i++)
        {
          SetLst.Add(string.Format("*Elset, elset={0}-{1}-TieSrf, internal, instance={1}", Tie.Branch(1)[i], ID));
          for (int j = 0; j < Inf.Branch(1, i, 4).Count; j++)
          {
            SetLst.Add(Inf.Branch(1, i, 4)[j]);
          }
        }
      }
      //Define tie surface;
      if ( Inf.PathExists(1, 0, 5) )
      {
        for (int i = 0; i < Tie.Branch(0).Count; i++)
        {
          SetLst.Add(string.Format("*Surface, type=ELEMENT, name={0}-{1}-TieSrf", Tie.Branch(1)[i], ID));
          SetLst.Add(string.Format("{0}-{1}-TieSrf, S{2}", Tie.Branch(1)[i], ID, Inf.Branch(1, i, 5)[0]));
        }
      }
      //Add elements on loaded face;
      if ( Inf.PathExists(1, 0, 2) )
      {
        for (int i = 0; i < Load.Branch(0).Count; i++)
        {
          SetLst.Add(string.Format("*Elset, elset={0}-Surf, internal, instance={1}", Load.Branch(2)[i], ID));
          for (int j = 0; j < Inf.Branch(1, i, 2).Count; j++)
          {
            SetLst.Add(Inf.Branch(1, i, 2)[j]);
          }
        }
      }
      //Define load surface;
      if ( Inf.PathExists(1, 0, 2) )
      {
        for (int i = 0; i < Load.Branch(0).Count; i++)
        {
          SetLst.Add(string.Format("*Surface, type=ELEMENT, name={0}", Load.Branch(2)[i]));
          SetLst.Add(string.Format("{0}-Surf, S{1}", Load.Branch(2)[i], Inf.Branch(1, i, 3)[0]));
        }
      }
      //Boundary conditions;
      if ( Support.PathExists(0) )
      {
        for (int i = 0; i < Support.Branch(1).Count; i++)
        {
          BCLst.Add(string.Format("{0}", Support.Branch(1)[i]));
        }
      }
      //Steps;
      if ( Load.PathExists(0) )
      {
        for (int i = 0; i < Load.Branch(1).Count; i++)
        {
          StpLst.Add(string.Format("{0}", Load.Branch(1)[i]));;
        }
      }

    }
    Geometry = GeoTree;
    Part = PrtLst;
    Instance = InsLst;
    Sets = SetLst;
    BC = BCLst;
    Steps = StpLst;
  }

  // <Custom additional code> 
  
  public int SrfInd(List < int > N)
  {
    int Indx;

    if (!N.Except(new List<int>{ 1, 2, 3, 4 }).Any() && N.Count == 4)
    {
      Indx = 1;
    }
    else if(!N.Except(new List<int>{ 5, 6, 7, 8 }).Any() && N.Count == 4)
    {
      Indx = 2;
    }
    else if(!N.Except(new List<int>{ 1, 2, 5, 6 }).Any() && N.Count == 4)
    {
      Indx = 3;
    }
    else if(!N.Except(new List<int>{ 2, 3, 6, 7 }).Any() && N.Count == 4)
    {
      Indx = 4;
    }
    else if(!N.Except(new List<int>{ 3, 4, 7, 8 }).Any() && N.Count == 4)
    {
      Indx = 5;
    }
    else
    {
      Indx = 6;
    }
    return Indx;
  }

  public bool IdenSet(Brep Srf, Point3d Nd, double Tol)
  {
    bool Truth;
    Point3d Pt = Srf.ClosestPoint(Nd);
    Vector3d Vct = Pt - Nd;

    if (Vct.Length < Tol)
    {
      Truth = true;
    }
    else
    {
      Truth = false;
    }
    return Truth;
  }

  public bool PtCmp(Point3d ElNd, Point3d Nd)
  {
    bool Truth;
    Vector3d Vct = ElNd - Nd;

    if (Vct.Length < doc.ModelAbsoluteTolerance)
    {
      Truth = true;
    }
    else
    {
      Truth = false;
    }
    return Truth;
  }
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

    DataTree<System.Object> Mesh = null;
    if (inputs[1] != null)
    {
      Mesh = GH_DirtyCaster.CastToTree<System.Object>(inputs[1]);
    }

    List<string> ElmntType = null;
    if (inputs[2] != null)
    {
      ElmntType = GH_DirtyCaster.CastToList<string>(inputs[2]);
    }
    List<System.Object> Material = null;
    if (inputs[3] != null)
    {
      Material = GH_DirtyCaster.CastToList<System.Object>(inputs[3]);
    }
    List<string> Orientation = null;
    if (inputs[4] != null)
    {
      Orientation = GH_DirtyCaster.CastToList<string>(inputs[4]);
    }
    DataTree<System.Object> Support = null;
    if (inputs[5] != null)
    {
      Support = GH_DirtyCaster.CastToTree<System.Object>(inputs[5]);
    }

    DataTree<System.Object> Load = null;
    if (inputs[6] != null)
    {
      Load = GH_DirtyCaster.CastToTree<System.Object>(inputs[6]);
    }

    DataTree<System.Object> Tie = null;
    if (inputs[7] != null)
    {
      Tie = GH_DirtyCaster.CastToTree<System.Object>(inputs[7]);
    }

    bool Run = default(bool);
    if (inputs[8] != null)
    {
      Run = (bool)(inputs[8]);
    }



    //3. Declare output parameters
      object Geometry = null;
  object Part = null;
  object Instance = null;
  object Sets = null;
  object BC = null;
  object Steps = null;


    //4. Invoke RunScript
    RunScript(ID, Mesh, ElmntType, Material, Orientation, Support, Load, Tie, Run, ref Geometry, ref Part, ref Instance, ref Sets, ref BC, ref Steps);
      
    try
    {
      //5. Assign output parameters to component...
            if (Geometry != null)
      {
        if (GH_Format.TreatAsCollection(Geometry))
        {
          IEnumerable __enum_Geometry = (IEnumerable)(Geometry);
          DA.SetDataList(0, __enum_Geometry);
        }
        else
        {
          if (Geometry is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(0, (Grasshopper.Kernel.Data.IGH_DataTree)(Geometry));
          }
          else
          {
            //assign direct
            DA.SetData(0, Geometry);
          }
        }
      }
      else
      {
        DA.SetData(0, null);
      }
      if (Part != null)
      {
        if (GH_Format.TreatAsCollection(Part))
        {
          IEnumerable __enum_Part = (IEnumerable)(Part);
          DA.SetDataList(1, __enum_Part);
        }
        else
        {
          if (Part is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(1, (Grasshopper.Kernel.Data.IGH_DataTree)(Part));
          }
          else
          {
            //assign direct
            DA.SetData(1, Part);
          }
        }
      }
      else
      {
        DA.SetData(1, null);
      }
      if (Instance != null)
      {
        if (GH_Format.TreatAsCollection(Instance))
        {
          IEnumerable __enum_Instance = (IEnumerable)(Instance);
          DA.SetDataList(2, __enum_Instance);
        }
        else
        {
          if (Instance is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(2, (Grasshopper.Kernel.Data.IGH_DataTree)(Instance));
          }
          else
          {
            //assign direct
            DA.SetData(2, Instance);
          }
        }
      }
      else
      {
        DA.SetData(2, null);
      }
      if (Sets != null)
      {
        if (GH_Format.TreatAsCollection(Sets))
        {
          IEnumerable __enum_Sets = (IEnumerable)(Sets);
          DA.SetDataList(3, __enum_Sets);
        }
        else
        {
          if (Sets is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(3, (Grasshopper.Kernel.Data.IGH_DataTree)(Sets));
          }
          else
          {
            //assign direct
            DA.SetData(3, Sets);
          }
        }
      }
      else
      {
        DA.SetData(3, null);
      }
      if (BC != null)
      {
        if (GH_Format.TreatAsCollection(BC))
        {
          IEnumerable __enum_BC = (IEnumerable)(BC);
          DA.SetDataList(4, __enum_BC);
        }
        else
        {
          if (BC is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(4, (Grasshopper.Kernel.Data.IGH_DataTree)(BC));
          }
          else
          {
            //assign direct
            DA.SetData(4, BC);
          }
        }
      }
      else
      {
        DA.SetData(4, null);
      }
      if (Steps != null)
      {
        if (GH_Format.TreatAsCollection(Steps))
        {
          IEnumerable __enum_Steps = (IEnumerable)(Steps);
          DA.SetDataList(5, __enum_Steps);
        }
        else
        {
          if (Steps is Grasshopper.Kernel.Data.IGH_DataTree)
          {
            //merge tree
            DA.SetDataTree(5, (Grasshopper.Kernel.Data.IGH_DataTree)(Steps));
          }
          else
          {
            //assign direct
            DA.SetData(5, Steps);
          }
        }
      }
      else
      {
        DA.SetData(5, null);
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