
using System;
using NXOpen;
using NXOpen.BlockStyler;
using System.Collections.Generic;
using System.Linq;
using NXOpen.Features;
using NXOpen.UF;
using NPOI;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.SS;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using System.Windows.Forms;
using System.Reflection;

public class PointManager
{
    //class members
    private static Session theSession = null;
    private static UI theUI = null;
    private static UFSession theUFSession = null;
    private string theDlxFileName;
    
    private NXOpen.BlockStyler.BlockDialog theDialog;
    private NXOpen.BlockStyler.Group group0;// Block type: Group
    private NXOpen.BlockStyler.Enumeration GraphTypeSelect;// Block type: Enumeration
    private NXOpen.BlockStyler.SuperPoint PointSelector;// Block type: Super Point
    private NXOpen.BlockStyler.Button ExportButton;// Block type: Button
    private NXOpen.BlockStyler.Group group;// Block type: Group
    private NXOpen.BlockStyler.Button ImportButton;// Block type: Button
    //------------------------------------------------------------------------------
    //Bit Option for Property: SnapPointTypesEnabled
    //------------------------------------------------------------------------------
    public static readonly int              SnapPointTypesEnabled_UserDefined = (1 << 0);
    public static readonly int                 SnapPointTypesEnabled_Inferred = (1 << 1);
    public static readonly int           SnapPointTypesEnabled_ScreenPosition = (1 << 2);
    public static readonly int                 SnapPointTypesEnabled_EndPoint = (1 << 3);
    public static readonly int                 SnapPointTypesEnabled_MidPoint = (1 << 4);
    public static readonly int             SnapPointTypesEnabled_ControlPoint = (1 << 5);
    public static readonly int             SnapPointTypesEnabled_Intersection = (1 << 6);
    public static readonly int                SnapPointTypesEnabled_ArcCenter = (1 << 7);
    public static readonly int            SnapPointTypesEnabled_QuadrantPoint = (1 << 8);
    public static readonly int            SnapPointTypesEnabled_ExistingPoint = (1 << 9);
    public static readonly int             SnapPointTypesEnabled_PointonCurve = (1 <<10);
    public static readonly int           SnapPointTypesEnabled_PointonSurface = (1 <<11);
    public static readonly int         SnapPointTypesEnabled_PointConstructor = (1 <<12);
    public static readonly int     SnapPointTypesEnabled_TwocurveIntersection = (1 <<13);
    public static readonly int             SnapPointTypesEnabled_TangentPoint = (1 <<14);
    public static readonly int                    SnapPointTypesEnabled_Poles = (1 <<15);
    public static readonly int         SnapPointTypesEnabled_BoundedGridPoint = (1 <<16);
    public static readonly int         SnapPointTypesEnabled_FacetVertexPoint = (1 <<17);
    //------------------------------------------------------------------------------
    //Bit Option for Property: SnapPointTypesOnByDefault
    //------------------------------------------------------------------------------
    public static readonly int             SnapPointTypesOnByDefault_EndPoint = (1 << 3);
    public static readonly int             SnapPointTypesOnByDefault_MidPoint = (1 << 4);
    public static readonly int         SnapPointTypesOnByDefault_ControlPoint = (1 << 5);
    public static readonly int         SnapPointTypesOnByDefault_Intersection = (1 << 6);
    public static readonly int            SnapPointTypesOnByDefault_ArcCenter = (1 << 7);
    public static readonly int        SnapPointTypesOnByDefault_QuadrantPoint = (1 << 8);
    public static readonly int        SnapPointTypesOnByDefault_ExistingPoint = (1 << 9);
    public static readonly int         SnapPointTypesOnByDefault_PointonCurve = (1 <<10);
    public static readonly int       SnapPointTypesOnByDefault_PointonSurface = (1 <<11);
    public static readonly int     SnapPointTypesOnByDefault_PointConstructor = (1 <<12);
    public static readonly int     SnapPointTypesOnByDefault_BoundedGridPoint = (1 <<16);
    //------------------------------------------------------------------------------
    //Bit Option for Property: EntityType
    //------------------------------------------------------------------------------
    public static readonly int                         EntityType_AllowPoints = (1 << 3);
    //------------------------------------------------------------------------------
    //Bit Option for Property: CurveRules
    //------------------------------------------------------------------------------
    public static readonly int                         CurveRules_SingleCurve = (1 << 0);
    public static readonly int                         CurveRules_InferCurves = (1 << 7);
    public static readonly int                       CurveRules_FeaturePoints = (1 <<10);
    
    //------------------------------------------------------------------------------
    //Constructor for NX Styler class
    //------------------------------------------------------------------------------
    public PointManager()
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            theUFSession = UFSession.GetUFSession();

            //theSession.ListingWindow.Open();
            //theSession.ListingWindow.WriteLine(AppDomain.CurrentDomain.BaseDirectory);

            theDlxFileName = AppDomain.CurrentDomain.BaseDirectory + @"\PointManager.dlx";
            theDialog = theUI.CreateDialog(theDlxFileName);
            theDialog.AddApplyHandler(new NXOpen.BlockStyler.BlockDialog.Apply(apply_cb));
            theDialog.AddOkHandler(new NXOpen.BlockStyler.BlockDialog.Ok(ok_cb));
            theDialog.AddUpdateHandler(new NXOpen.BlockStyler.BlockDialog.Update(update_cb));
            theDialog.AddInitializeHandler(new NXOpen.BlockStyler.BlockDialog.Initialize(initialize_cb));
            theDialog.AddDialogShownHandler(new NXOpen.BlockStyler.BlockDialog.DialogShown(dialogShown_cb));
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            throw ex;
        }
    }

    public static void Main()
    {
        PointManager thePointManager = null;
        try
        {
            thePointManager = new PointManager();
            // The following method shows the dialog immediately
            thePointManager.Show();
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        finally
        {
            if (thePointManager != null)
                thePointManager.Dispose();
            thePointManager = null;
        }
    }

     public static int GetUnloadOption(string arg)
    {
        //return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);
         return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        // return System.Convert.ToInt32(Session.LibraryUnloadOption.AtTermination);
    }
    

    public static void UnloadLibrary(string arg)
    {
        try
        {
            //---- Enter your code here -----
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }
    
    //------------------------------------------------------------------------------
    //This method shows the dialog on the screen
    //------------------------------------------------------------------------------
    public NXOpen.UIStyler.DialogResponse Show()
    {
        try
        {
            theDialog.Show();
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        return 0;
    }
    
    //------------------------------------------------------------------------------
    //Method Name: Dispose
    //------------------------------------------------------------------------------
    public void Dispose()
    {
        if(theDialog != null)
        {
            theDialog.Dispose();
            theDialog = null;
        }
    }
    
    //------------------------------------------------------------------------------
    //---------------------Block UI Styler Callback Functions--------------------------
    //------------------------------------------------------------------------------
    
    //------------------------------------------------------------------------------
    //Callback Name: initialize_cb
    //------------------------------------------------------------------------------
    public void initialize_cb()
    {
        try
        {
            group0 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group0");
            GraphTypeSelect = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("GraphTypeSelect");
            PointSelector = (NXOpen.BlockStyler.SuperPoint)theDialog.TopBlock.FindBlock("PointSelector");
            ExportButton = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("ExportButton");
            group = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group");
            ImportButton = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("ImportButton");
            string[] types = { "Gap or Size", "Curvature", "Random Points" };
            GraphTypeSelect.SetEnumMembers(types);

            PointSelector.StepStatusAsString = "Optional";

            ExportButton.Enable = false;
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }
    
    //------------------------------------------------------------------------------
    //Callback Name: dialogShown_cb
    //This callback is executed just before the dialog launch. Thus any value set 
    //here will take precedence and dialog will be launched showing that value. 
    //------------------------------------------------------------------------------
    public void dialogShown_cb()
    {
        try
        {
            //---- Enter your callback code here -----
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }
    
    //------------------------------------------------------------------------------
    //Callback Name: apply_cb
    //------------------------------------------------------------------------------
    public int apply_cb()
    {
        int errorCode = 0;
        try
        {
            //---- Enter your callback code here -----
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            errorCode = 1;
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        return errorCode;
    }
    //

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
        public NodePoint(string name, double Xpos, double Ypos, double Zpos)
        {
            Name = name;
            X = Xpos;
            Y = Ypos;
            Z = Zpos;
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

    private LinkedList<NodePoint> Graph = new LinkedList<NodePoint>();

    private List<NodePoint> AllPoints = new List<NodePoint>(); 

    private void ComposeXLSX()
    {
        XSSFWorkbook outbook = new XSSFWorkbook();
        outbook.CreateSheet("Coordinates");
        var Sheet = outbook.GetSheetAt(0);
        var Row = Sheet.CreateRow(0);
        ICell Cell = Row.CreateCell(0);
        Cell.SetCellValue("Name");
        Cell = Row.CreateCell(1);
        Cell.SetCellValue("X");
        Cell = Row.CreateCell(2);
        Cell.SetCellValue("Y");
        Cell = Row.CreateCell(3);
        Cell.SetCellValue("Z");
        int rnum = 1;
        if (Graph.Count == 0)
        {
            foreach (NodePoint node in AllPoints)
            {
                Row = Sheet.CreateRow(rnum++);
                Cell = Row.CreateCell(0);
                Cell.SetCellType(CellType.String);
                Cell.SetCellValue(node.Name);
                Cell = Row.CreateCell(1);
                Cell.SetCellValue(Math.Round(node.X,3));
                Cell = Row.CreateCell(2);
                Cell.SetCellValue(Math.Round(node.Y, 3));
                Cell = Row.CreateCell(3);
                Cell.SetCellValue(Math.Round(node.Z, 3));
            }
        }
        else
        {
            for (LinkedListNode<NodePoint> node = Graph.First; node != null; node = node.Next)
            {
                Row = Sheet.CreateRow(rnum++);
                Cell = Row.CreateCell(0);
                Cell.SetCellType(CellType.String);
                Cell.SetCellValue(node.Value.Name);
                Cell = Row.CreateCell(1);
                Cell.SetCellValue(Math.Round(node.Value.X, 3));
                Cell = Row.CreateCell(2);
                Cell.SetCellValue(Math.Round(node.Value.Y, 3));
                Cell = Row.CreateCell(3);
                Cell.SetCellValue(Math.Round(node.Value.Z, 3));
            }
        }


        var myfont = outbook.CreateFont();
        myfont.FontHeightInPoints = 11;
        myfont.FontName = "Times New Roman";
        var myStyle = outbook.CreateCellStyle();
        myStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        myStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        myStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        myStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        myStyle.SetFont(myfont);
        myStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        myStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

        for (int i = 0; i < Sheet.PhysicalNumberOfRows; i++)
        {
            Row = Sheet.GetRow(i);
            for (int j =0 ; j <= 3; j++)
            {
                Cell = Row.GetCell(j);
                Cell.CellStyle = myStyle;
            }
        }


        SaveFileDialog myDialog = new SaveFileDialog();
        myDialog.Filter = "Документы Excel (*.xlsx)|**.XLSX";
        string f ="";
        theUFSession.Part.AskPartName(theSession.Parts.Work.Tag, out f);
        myDialog.FileName = f + "_" + "XYZ" + ".xlsx";
        if (myDialog.ShowDialog() == DialogResult.OK)
        {
            var FileName = myDialog.FileName;
            var outFile = new FileStream(FileName, FileMode.Create, FileAccess.Write);
            outbook.Write(outFile);
            theUI.NXMessageBox.Show("Завершено", NXMessageBox.DialogType.Information, "Файл сохранен.");
        }
    }

    private  HSSFWorkbook OpenFileXLS(string FileName)
    {
        try
        {
            var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            var Doc = new HSSFWorkbook(inputFile);
            return Doc;
        }
        catch (IOException)
        {
            return null;
        }
    }

    private  XSSFWorkbook OpenFileXLSX(string FileName)
    {
        try
        {
            var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            var Doc = new XSSFWorkbook(inputFile);
            return Doc;
        }
        catch (IOException)
        {
            return null;
        }

    }

    private void DrawPoints(IWorkbook Doc)
    {
        AllPoints.Clear();
        int errcount = 0;
        ISheet Sheet = Doc.GetSheetAt(0);
        for (int i = 1; i < Sheet.PhysicalNumberOfRows; i++)
        {
            if (Sheet.GetRow(i) != null)
            {
                try
                {
                    IRow Row = Sheet.GetRow(i);
                    string name = Row.GetCell(0).ToString();
                    double X = Convert.ToDouble(Row.GetCell(1).ToString().Replace('.', ','));
                    double Y = Convert.ToDouble(Row.GetCell(2).ToString().Replace('.', ','));
                    double Z = Convert.ToDouble(Row.GetCell(3).ToString().Replace('.', ','));
                    NodePoint iNode = new NodePoint(name, X, Y, Z);
                    AllPoints.Add(iNode);
                }
                catch
                {
                    errcount++;
                }
            }
        }

        int counter = 0;
        var part = theSession.Parts.Work;
        var Tags = new List<Tag>();
        var tag = new Tag();
        List<NXObject> Tlist = new List<NXObject>();
        foreach (NodePoint node in AllPoints)
        {
            var newpoint = new Point3d(node.X, node.Y, node.Z);
            var p = part.Points.CreatePoint(newpoint);
            Tlist.Add(p as NXObject);
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
        theUFSession.Modl.CreateSetOfFeature("Points", Tags.ToArray(), Tags.Count, 1, out tag);
        theUI.NXMessageBox.Show("Завершено", NXMessageBox.DialogType.Information, "Ошибки: " + errcount.ToString());
    }

    //------------------------------------------------------------------------------
    //Callback Name: update_cb
    //------------------------------------------------------------------------------
    public int update_cb( NXOpen.BlockStyler.UIBlock block)
    {

        try
        {
            if (block == GraphTypeSelect)
            {

            }
            else if(block == PointSelector)
            {
                AllPoints.Clear();
                int counter = 0;
                var Selection = PointSelector.GetSelectedObjects();
                var Sect = Selection[0] as Section;
                NXObject[] Data;
                Sect.EvaluateAndAskOutputEntities(out Data);
                foreach (NXObject obj in Data)
                {
                    Point3d coords = ((Point)obj).Coordinates;
                    NodePoint iNode = new NodePoint("Point_" + counter++, coords);
                    if (!AllPoints.Any(p => Math.Round(p.X,3)  == Math.Round(coords.X,3) && Math.Round(p.Y, 3) == Math.Round(coords.Y, 3) && Math.Round(p.Z, 3) == Math.Round(coords.Z, 3)))
                    {
                        AllPoints.Add(iNode); // Собираем список всех точек в NodeList
                    }
                }
                if (AllPoints.Count != 0)
                {
                    ExportButton.Enable = true;
                }
                else
                {
                    ExportButton.Enable = false;
                }
            }

            else if(block == ExportButton)
            {
                var type = GraphTypeSelect.ValueAsString;
                var part = theSession.Parts.Work;
                var Tags = new List<Tag>();
                var tag = new Tag();
                int counter = 0;
                switch (type)
                {
                    case "Gap or Size":
                        #region Gap Calculation
                            
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
                        Graph.Clear();
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
                        theUFSession.Modl.CreateSetOfFeature("Contour Points", Tags.ToArray(), Tags.Count, 1, out tag);

                        break;
                    #endregion
                    case "Curvature":
                        #region Curvature Calculation
                        var Zsort = new SortByZ();
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
                        theUFSession.Modl.CreateSetOfFeature("Curvature", Tags.ToArray(), Tags.Count, 1, out tag);
                        break;
                    #endregion
                    case "Random Points":
                        #region Random Calculation
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
                        theUFSession.Modl.CreateSetOfFeature("Points", Tags.ToArray(), Tags.Count, 1, out tag);
                        break;
                        #endregion
                }

                ComposeXLSX();

            }
            else if(block == ImportButton)
            {
                OpenFileDialog myDialog = new OpenFileDialog();
                string FileName = "";
                myDialog.Filter = "Документы Excel (*.xlsx)|**.XLSX";
                if (myDialog.ShowDialog() == DialogResult.OK)
                {
                    FileName = myDialog.FileName;
                    var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                    if (FileName.ToLower().EndsWith(".xlsx"))
                    {
                       XSSFWorkbook inbook =  OpenFileXLSX(FileName);
                        DrawPoints(inbook);
                    }
                    if (FileName.ToLower().EndsWith(".xls"))
                    {
                       HSSFWorkbook inbook = OpenFileXLS(FileName);
                        DrawPoints(inbook);
                    }
                }

 
            }
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        return 0;
    }
    
    //------------------------------------------------------------------------------
    //Callback Name: ok_cb
    //------------------------------------------------------------------------------
    public int ok_cb()
    {
        int errorCode = 0;
        try
        {
            errorCode = apply_cb();
            //---- Enter your callback code here -----
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            errorCode = 1;
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        return errorCode;
    }
    
    //------------------------------------------------------------------------------
    //Function Name: GetBlockProperties
    //Returns the propertylist of the specified BlockID
    //------------------------------------------------------------------------------
    public PropertyList GetBlockProperties(string blockID)
    {
        PropertyList plist =null;
        try
        {
            plist = theDialog.GetBlockProperties(blockID);
        }
        catch (Exception ex)
        {
            //---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
        }
        return plist;
    }
    
}
