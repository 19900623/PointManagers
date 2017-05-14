
using System;
using NXOpen;
using System.Windows.Forms;
using NXOpenUI;
using System.IO;
using NXOpen.UF;
using NPOI;
using System.Collections.Generic;
using System.Linq;
using NXOpen.Features;
using NXOpen.BlockStyler;

public class Program
{
    // class members
    private static Session theSession;
    private static UFSession uf;
    private static UI theUI;

    class NodePoint
    {
        public string Name { get; set; }
        public double X;
        public double Y;
        public double Z;
        public NodePoint(string name, Point3d coords)
        {
            Name = name;
            X = coords.X;
            Y = coords.Y;
            Z = coords.Z;
        }
        public Dictionary<NodePoint, double> DistanceTo = new Dictionary<NodePoint, double>();

        public NodePoint getClosestNode()
        {
            NodePoint result = null;
            double min = DistanceTo.Values.Min();
            foreach (NodePoint n in DistanceTo.Keys)
            {
                if (DistanceTo[n] <= min)
                {
                    result = n;
                }
            }
            return result;
        }
    }

    class SortByZ : IComparer<NodePoint>
    {
        public int Compare(NodePoint A, NodePoint B) 
        {
            if (A.Z < B.Z) { return 1; } 
            if (A.Z > B.Z) { return -1; } 
            return 0;
        }
    }

    private static double getDistance(NodePoint A, NodePoint B)
    {
        double dx = B.X - A.X;
        double dy = B.Y - A.Y;
        double dz = B.Z - A.Z;
        return Math.Abs((Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz))));
    }

    private static LinkedList<NodePoint> Graph = new LinkedList<NodePoint>();


    public static int NotMain(string[] args)
    {

        
        theSession = Session.GetSession();
        uf = UFSession.GetUFSession();
        theUI = UI.GetUI();        
        var Zsort = new SortByZ();
        var part = theSession.Parts.Work;
        var AllPoints = new List<NodePoint>();  
        int counter = 0;

        for (int i=0; i < theUI.SelectionManager.GetNumSelectedObjects(); i++)
        {
            TaggedObject obj = theUI.SelectionManager.GetSelectedTaggedObject(i);
            if (obj is Point)
            {
                Point3d coords = ((Point)obj).Coordinates;
                NodePoint iNode = new NodePoint("Point_" + counter++, coords);
                AllPoints.Add(iNode); // Собираем список всех точек
            }
        }
        
        var Tags = new List<Tag>();
        var tag = new Tag();
        
        /* // Curvature = OK
        AllPoints.Sort(Zsort);
        foreach (NodePoint node in AllPoints)
        {
            var newpoint = new Point3d(node.X, node.Y, node.Z);
            var p = part.Points.CreatePoint(newpoint);
            node.Name = "A" + counter++;
            p.SetName(node.Name);
            PointFeatureBuilder p_feature;
            p_feature = part.BaseFeatures.CreatePointFeatureBuilder(null);
            p_feature.Point = p;
            p_feature.Commit();
            p_feature.GetFeature().SetName(node.Name);
            Tags.Add(p_feature.GetFeature().Tag);
            p.SetVisibility(SmartObject.VisibilityOption.Visible);
        }
        uf.Modl.CreateSetOfFeature("Curvature", Tags.ToArray(), Tags.Count, 1, out tag);
        */


        /*
        foreach (NodePoint Vstart in AllPoints)
        {
            foreach (NodePoint Vend in AllPoints)
            {
                if (Vstart.Name != Vend.Name)
                {
                    Vstart.DistanceTo.Add(Vend, getDistance(Vstart, Vend));
                }
            }
        }

        
        NodePoint thisNode = AllPoints[0];
        NodePoint nextNode = null;
        counter = 0;
        thisNode.Name = "A" + counter++;
        Graph.AddLast(thisNode);

        while (Graph.Count < AllPoints.Count)
        {
            nextNode = thisNode.getClosestNode();
            while (Graph.Contains(nextNode))
            {
                thisNode.DistanceTo.Remove(nextNode);
                nextNode = thisNode.getClosestNode();
            }
            thisNode = nextNode;
            thisNode.Name = "A" + counter++;
            Graph.AddLast(thisNode);
        }

        for (LinkedListNode<NodePoint> node = Graph.First; node != null; node = node.Next)
        {
            var newpoint = new Point3d(node.Value.X, node.Value.Y, node.Value.Z);
            var p = part.Points.CreatePoint(newpoint);
            p.SetName(node.Value.Name);
            PointFeatureBuilder p_feature;
            p_feature = part.BaseFeatures.CreatePointFeatureBuilder(null);
            p_feature.Point = p;
            p_feature.Commit();
            p_feature.GetFeature().SetName(node.Value.Name);
            Tags.Add(p_feature.GetFeature().Tag);
            p.SetVisibility(SmartObject.VisibilityOption.Visible);
        }


        uf.Modl.CreateSetOfFeature("Curvature", Tags.ToArray(), Tags.Count, 1, out tag);
        */
        return 0;
    }

    public static int GetUnloadOption(string arg)
    {

        return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);

    }
}

