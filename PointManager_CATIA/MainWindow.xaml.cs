using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using INFITF;
using MECMOD;
using HybridShapeTypeLib;
using NPOI;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.SS;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using System.Reflection;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using AnnotationTypeLib;
using System.ComponentModel;

namespace PointManager_CATIA
{
    public class NodePoint
    {
        public string Name { get; set; }
        public double X;
        public double Y;
        public double Z;
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

    public class SortByZ : IComparer<NodePoint>
    {
        public int Compare(NodePoint A, NodePoint B)
        {
            if (A.Z < B.Z) { return 1; }
            if (A.Z > B.Z) { return -1; }
            return 0;
        }
    }

    public partial class MainWindow : System.Windows.Window
    {
        public static INFITF.Application CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");

        private static double getDistance(NodePoint A, NodePoint B)
        {
            double dx = B.X - A.X;
            double dy = B.Y - A.Y;
            double dz = B.Z - A.Z;
            return Math.Abs((Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz))));
        }

        private LinkedList<NodePoint> Graph = new LinkedList<NodePoint>();

        private List<NodePoint> AllPoints = new List<NodePoint>();

        private HSSFWorkbook OpenFileXLS(string FileName)
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

        private XSSFWorkbook OpenFileXLSX(string FileName)
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
                    Cell.SetCellValue(Math.Round(node.X, 3));
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
                for (int j = 0; j <= 3; j++)
                {
                    Cell = Row.GetCell(j);
                    Cell.CellStyle = myStyle;
                }
            }


            SaveFileDialog myDialog = new SaveFileDialog();
            myDialog.Filter = "Документы Excel (*.xlsx)|**.XLSX";
            string f = "";
            f = CATIA.ActiveDocument.get_Name();
            myDialog.FileName = f + "_" + "XYZ" + ".xlsx";
            if (myDialog.ShowDialog() == true)
            {
                var FileName = myDialog.FileName;
                var outFile = new FileStream(FileName, FileMode.Create, FileAccess.Write);
                outbook.Write(outFile);
                var result = MessageBox.Show("Файл сохранен", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
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
            ////
            var partDoc = CATIA.ActiveDocument as PartDocument;
            var Part = partDoc.Part;
            var HBodies = Part.HybridBodies;
            var HBody = HBodies.Add();
            HBody.set_Name("Import Points_" + AllPoints.Count.ToString());
            Part.Update();
            var ShapeFactory = Part.HybridShapeFactory as HybridShapeFactory;
            foreach (NodePoint iNode in AllPoints)
            {
                var Point = ShapeFactory.AddNewPointCoord(iNode.X, iNode.Y, iNode.Z);
                Point.set_Name(iNode.Name);
                HBody.AppendHybridShape(Point);
            }
            Part.Update();
        }


        private void SetAnnotations(string Type)
        {
            var partDoc = CATIA.ActiveDocument as PartDocument;
            var Part = partDoc.Part;
            
            var AnnSets = Part.AnnotationSets as AnnotationSets;
            var AnnSet = AnnSets.Add("ISO_3D") as AnnotationSet;
            var AnnFactory = AnnSet.AnnotationFactory as AnnotationFactory;
            var HBodies = Part.HybridBodies;
            var HBody = HBodies.Item("Import Points_" + AllPoints.Count.ToString());
            var HShapes = HBody.HybridShapes as HybridShapes;
            var Surfaces = Part.UserSurfaces as UserSurfaces;
            double Z0 = AllPoints[0].Z;
            foreach (NodePoint iNode in AllPoints)
            {
                var Point = HShapes.Item(iNode.Name) as HybridShapePointCoord;
                var Ref = Part.CreateReferenceFromObject(Point) as Reference;
                var Surf = Surfaces.Generate(Ref);

                if (Type == "WS")
                {
                    var Ann = AnnFactory.CreateEvoluateText(Surf, iNode.Y + 15, iNode.X + 15, Z0 - iNode.Z, true);
                    Ann.Text().set_Text(iNode.Name);
                }
                if (Type == "BL")
                {
                    var Ann = AnnFactory.CreateEvoluateText(Surf, -iNode.Y + 15, iNode.X + 15, Z0 - iNode.Z, true);
                    Ann.Text().set_Text(iNode.Name);
                }
                if (Type == "LD")
                {
                    var Ann = AnnFactory.CreateEvoluateText(Surf, iNode.X + 15, iNode.Y + 15, Z0 - iNode.Z, true);
                    Ann.Text().set_Text(iNode.Name);
                }
                if (Type == "RD")
                {
                    var Ann = AnnFactory.CreateEvoluateText(Surf, -iNode.X + 15, -iNode.Y + 15, Z0 - iNode.Z, true);
                    Ann.Text().set_Text(iNode.Name);
                }
            }
            Part.Update();
            var result = MessageBox.Show("Готово!", "Точки проставлены.", MessageBoxButton.OK, MessageBoxImage.Information);

        }

        public MainWindow()
        {
            InitializeComponent();

        }

        private void GraphTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            AllPoints.Clear();
            Selection sel = CATIA.ActiveDocument.Selection;
            sel.Search("((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),sel");
            //sel.Search("CATPrtsearch.point, sel");
            labelCount.Content = "Точек выбрано: " + sel.Count2.ToString();
            object[] coords = new object[3];
            int counter = 0;

            if (sel.Count != 0)
            {
                for (int i = 1; i <= sel.Count2; i++)
                {
                    var item = sel.Item2(i).Value as HybridShapeTypeLib.Point;
                    item.GetCoordinates(coords);
                    double X = Convert.ToDouble(coords[0]);
                    double Y = Convert.ToDouble(coords[1]);
                    double Z = Convert.ToDouble(coords[2]);
                    NodePoint iNode = new NodePoint("Point_" + counter++, X, Y, Z);
                    if (!AllPoints.Any(p => Math.Round(p.X, 3) == Math.Round(X, 3) && Math.Round(p.Y, 3) == Math.Round(Y, 3) && Math.Round(p.Z, 3) == Math.Round(Z, 3)))
                    {
                        AllPoints.Add(iNode); // Собираем список всех точек в NodeList
                    }
                }
                //
                ComposeXLSX();
                
            }
            else
            {
                  MessageBox.Show("Ни одной точки не выбрано!" + Environment.NewLine + "Всегда выбирайте точки с помощью рамки выделения, либо выбирая их названия из списка деталей.", "Упс!", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }



        }


        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog myDialog = new OpenFileDialog();
            string FileName = "";
            myDialog.Filter = "Документы Excel (*.xlsx)|**.XLSX";
            if (myDialog.ShowDialog() == true)
            {
                FileName = myDialog.FileName;
                var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                if (FileName.ToLower().EndsWith(".xlsx"))
                {
                    XSSFWorkbook inbook = OpenFileXLSX(FileName);
                    DrawPoints(inbook);
                    SetAnnotations("BL");
                    
                }
                if (FileName.ToLower().EndsWith(".xls"))
                {
                    HSSFWorkbook inbook = OpenFileXLS(FileName);
                    DrawPoints(inbook);
                    SetAnnotations("BL");
                }
            }


        }

        private void ViewTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
