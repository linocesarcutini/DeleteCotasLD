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

                doc.SendCommand("EL");

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