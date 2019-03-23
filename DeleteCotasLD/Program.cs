using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace ChangeTextInVariousDrawings
{

    class Program
    {

        private static string[] arquivos;

        [STAThread]
        public static void Main(string[] args)
        {
            if (args.Length != 0)
            {
                arquivos = args;
            }
            else
            {
                SelectFiles();
            }

            AcadApplication acApp = null;

            try
            {
                acApp = Marshal.GetActiveObject("AutoCAD.Application") as AcadApplication;
            }
            catch
            {
                MessageBox.Show("Não foi possível abrir o AutoCAD");
            }

            foreach (var desenho in arquivos)
            {
                AcadDocument doc;

                try
                {
                    doc = acApp.Documents.Open(desenho, false, string.Empty);

                    AcadModelSpace modelSpace = doc.ModelSpace;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Não foi possível abrir o desenho " + desenho + ex);
                    break;
                }

                doc.SetVariable("CLAYER", "0");

                AcadSelectionSet selset = null;
                selset = doc.SelectionSets.Add("layer");
                short[] ftype = { 8 };
                string[] fdataX = {
                    "COTAS", "L.D.", "L1", "L2 ", "L3", "L4", "L5",
                    "REV.L1", "REV.L2", "REV.L3","REV.L4","REV.L5",
                    "00L1","00L2","00L3","00L4","00L5",
                    "R00L1","R00L2","R00L3","R00L4","R00L5",
                    "0AL1","0AL2","0AL3","0AL4","0AL5",
                    "0BL1","0BL2","0BL3","0BL4","0BL5",
                    "0CL1","0CL2","0CL3","0CL4","0CL5",
                    "01L1","01L2","01L3","01L4","01L5",
                    "02L1","02L2","02L3","02L4","02L5",
                    "R01L1","R01L2","R01L3","R01L4","R01L5",
                    "A1L1","A1L2","A1L3","A1L4","A1L5", "RL01",
                    "RL001", "R0011", "RL2", "RL3", "RL4"};

                foreach (var item in fdataX)
                {
                    object[] fdata = { item };

                    selset.Select(AcSelect.acSelectionSetAll, null, null, ftype, fdata);

                    if (selset.Count != 0)
                    {
                        selset.Erase();
                    }
                }

                doc.PurgeAll();
                doc.PurgeAll();
                doc.PurgeAll();
                acApp.ZoomExtents();
                doc.Save();
                doc.Close();
            }
        }

        static void SelectFiles()
        {

            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Drawing Files|*.dwg";
            openFileDialog2.Title = "Selecione os arquivos DWG";
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.Multiselect = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                // Guarda a lista dos caminhos completos dos arquivos selecionados
                arquivos = openFileDialog2.FileNames;
            }
            else
            {
                Console.WriteLine("Arquivos não selecionados!");
                Console.ReadKey();
            }
        }
    }
}