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
using System.Text.RegularExpressions;

namespace PointManager_CATIA
{

    public partial class MainWindow : System.Windows.Window
    {

        public class NodePoint
        {
            public string Name { get; set; }
            public double X;
            public double Y;
            public double Z;
            public NodePoint(string name, double Xpos, double Ypos, double Zpos, HybridShapeTypeLib.Point LinkToPoint)
            {
                Name = name;
                X = Xpos;
                Y = Ypos;
                Z = Zpos;
                link = LinkToPoint;
            }
            public HybridShapeTypeLib.Point link = null;

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

        private static double getDistance(NodePoint A, NodePoint B)
        {
            double dx = B.X - A.X;
            double dy = B.Y - A.Y;
            double dz = B.Z - A.Z;
            return Math.Abs((Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz))));
        }

        private LinkedList<NodePoint> Graph = new LinkedList<NodePoint>();

        private List<NodePoint> AllPoints = new List<NodePoint>();

        public class SortByNameDigit : IComparer<NodePoint>
        {
            public int Compare(NodePoint A, NodePoint B)
            {
                var A_Digitstring = Regex.Match(A.Name, @"\d+").Value;
                var A_num = Int32.Parse(A_Digitstring);
                var B_Digitstring = Regex.Match(B.Name, @"\d+").Value;
                var B_num = Int32.Parse(B_Digitstring);
                if (A_num > B_num) { return 1; }
                if (A_num < B_num) { return -1; }
                return 0;
            }
        }


        BackgroundWorker ExportWorker;

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            ExportWorker = new BackgroundWorker();
            ExportWorker.DoWork += ExportWorker_DoWork;
            ExportWorker.RunWorkerCompleted += ExportWorker_RunWorkerCompleted;
            ExportWorker.RunWorkerAsync();
            ExportButton.Content = "Подождите...";
            ImportButton.Content = "Подождите...";
            ImportButton.IsEnabled = false;
            ExportButton.IsEnabled = false;
        }

        private void ExportWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            AllPoints.Clear();
            CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");
            Selection sel = CATIA.ActiveDocument.Selection;
            sel.Search("((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),sel");
            //sel.Search("CATPrtsearch.point, sel");
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
                    string name = "";
                    if (item.get_Name().Contains("Точка.") || item.get_Name().Contains("Point."))
                    {
                        name = "Point_" + counter++;
                    }
                    else
                    {
                        name = item.get_Name();
                    }
                    NodePoint iNode = new NodePoint(name, X, Y, Z, item);
                    if (!AllPoints.Any(p => Math.Round(p.X, 3) == Math.Round(X, 3) && Math.Round(p.Y, 3) == Math.Round(Y, 3) && Math.Round(p.Z, 3) == Math.Round(Z, 3)))
                    {
                        AllPoints.Add(iNode); // Собираем список всех точек в NodeList
                    }
                }
                /////

                if (AllPoints.All(x => x.Name.Contains("Point_"))) //Если у все хточек нет нормальных названий
                {
                    var type = GraphTypeCombo.SelectedValue.ToString();
                    //var type = "Random Points";
                    counter = 0;
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
                                node.Value.link.set_Name(node.Value.Name);
                            }
                            break;
                        #endregion
                        case "Curvature":
                            #region Curvature Calculation
                            var Zsort = new SortByZ();
                            AllPoints.Sort(Zsort);
                            foreach (NodePoint node in AllPoints)
                            {
                                node.Name = "A" + counter++;
                                node.link.set_Name(node.Name);
                            }
                            break;
                        #endregion
                        case "Random Points":
                            #region Random Calculation
                            foreach (NodePoint node in AllPoints)
                            {
                                node.Name = "A" + counter++;
                                node.link.set_Name(node.Name);
                            }
                            break;
                            #endregion
                    }
                }
                else
                {
                    var Dsort = new SortByNameDigit();
                    AllPoints.Sort(Dsort);
                }
                //
                ComposeXLSX();
                var result = 
                e.Result = "ok";
            }
            else
            {
                e.Result = "no points";
            }
        }


        private void ExportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var strres = e.Result as string;
            ExportButton.Content = "Выполнить";
            ImportButton.Content = "Открыть файл";
            ImportButton.IsEnabled = true;
            ExportButton.IsEnabled = true;

            if (strres == "ok")
            {
                MessageBox.Show("Файл сохранен", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            if (strres == "no points")
            {
                MessageBox.Show("Ни одной точки не выбрано!" + Environment.NewLine + "Всегда выбирайте точки с помощью рамки выделения, либо выбирая их названия из списка деталей.", "Упс!", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                
            }
        }

    }
}

