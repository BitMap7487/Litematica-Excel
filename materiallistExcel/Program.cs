using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace materiallistExcel
{
    class Program
    {

        [STAThread]
        static void Main(string[] args)
        {

            Console.SetWindowSize(130, 30);

            Console.Clear();

            Console.WriteLine();
            Console.Title = "[Litematica Excel] by BitMap";
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine();
            foreach (string str in new List<string>()
            {

                " ██╗     ██╗████████╗███████╗███╗   ███╗ █████╗ ████████╗██╗ ██████╗ █████╗     ███████╗██╗  ██╗ ██████╗███████╗██╗     ",
                " ██║     ██║╚══██╔══╝██╔════╝████╗ ████║██╔══██╗╚══██╔══╝██║██╔════╝██╔══██╗    ██╔════╝╚██╗██╔╝██╔════╝██╔════╝██║     ",
                " ██║     ██║   ██║   █████╗  ██╔████╔██║███████║   ██║   ██║██║     ███████║    █████╗   ╚███╔╝ ██║     █████╗  ██║     ",
                " ██║     ██║   ██║   ██╔══╝  ██║╚██╔╝██║██╔══██║   ██║   ██║██║     ██╔══██║    ██╔══╝   ██╔██╗ ██║     ██╔══╝  ██║     ",
                " ███████╗██║   ██║   ███████╗██║ ╚═╝ ██║██║  ██║   ██║   ██║╚██████╗██║  ██║    ███████╗██╔╝ ██╗╚██████╗███████╗███████╗",
                " ╚══════╝╚═╝   ╚═╝   ╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝   ╚═╝   ╚═╝ ╚═════╝╚═╝  ╚═╝    ╚══════╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝",



                "\n",
                "╔════════════════════════════════════════════════════════════════════════════════════╗",
                "║                                Made by: BitMap#7487                                ║",
                "║                                                                                    ║",
                "║               Change the ugly material list.txt to an excel variant.               ║",
                "╚════════════════════════════════════════════════════════════════════════════════════╝",
                "\n",
                "\n"

            }) Console.WriteLine(string.Format("{0," + (object)(Console.WindowWidth / 2 + str.Length / 2) + "}", (object)str));
            
            Console.ResetColor();


            Console.WriteLine("\n\nPress enter to select the file.");
            Console.ReadLine();

            SelectFileTXT();
            RemoveSpacesAndFirstLines();
            ParseLines();
            TxtToExcel();

            string NewPath = $"Material List {DateTime.Now}.xlsx";

            Console.WriteLine($"All Done. File can be found in this programs folder with the name: {string.Join("-", NewPath.Split(Path.GetInvalidFileNameChars()))}");

        }

        static string TXTfilePath = string.Empty;
        static string FileContent = string.Empty;

        [STAThread]
        static void SelectFileTXT()
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {

                    TXTfilePath = openFileDialog.FileName;

                    using (StreamReader file = new StreamReader(TXTfilePath))
                    {
                        FileContent = file.ReadToEnd();
                    }

                }
            }

        }

        static void RemoveSpacesAndFirstLines()
        {

            Console.WriteLine("Removing all spaces and lines:");

            int iChar = 0;
            while (FileContent[iChar] != '|')
            {
                iChar++;
            }

            FileContent = FileContent.Remove(FileContent.Length - iChar * 3);

            var builder = new StringBuilder();

            foreach (char character in FileContent)
            {

                if (character != ' ')
                {
                    builder.Append(character);
                }

            }


            builder.Remove(0, builder.ToString().LastIndexOf('+') + 3);

            FileContent = builder.ToString();

        }

        public static List<string[]> ParsedData = new List<string[]>();

        static void ParseLines()
        {

            Console.WriteLine("Parsing Data:");

            foreach (string line in FileContent.Split('\n'))
            {

                string[] splitLine = line.Split('|');

                /*
                Item = splitLine[1]
                Total = splitLine[2]
                Missing = splitLine[3]
                Available = splitLine[4]
                */

                string[] parsed = { splitLine[1], splitLine[2], splitLine[3], splitLine[4] };

                ParsedData.Add(parsed);

            }

        }

        static void TxtToExcel()
        {

            Console.WriteLine("Generating XLSX file:");

            ExcelPackage excel = new ExcelPackage();

            var workSheet = excel.Workbook.Worksheets.Add("Material_List");

            workSheet.Cells[$"A1"].Value = "Item";
            workSheet.Cells[$"B1"].Value = "Total";
            workSheet.Cells[$"C1"].Value = "Missing";
            workSheet.Cells[$"D1"].Value = "Available";

            for (int i = 0; i <= ParsedData.Count - 1; i++)
            {
                string[] data = ParsedData[i];

                workSheet.Cells[$"A{i + 2}"].Value = data[0];
                workSheet.Cells[$"B{i + 2}"].Value = data[1];
                workSheet.Cells[$"C{i + 2}"].Formula = $"B{i+2}-D{i+2}";
                workSheet.Cells[$"D{i + 2}"].Value = data[3];

            }

            string NewPath = $"Material List {DateTime.Now}.xlsx";

            excel.SaveAs(string.Join("-", NewPath.Split(Path.GetInvalidFileNameChars())));

        }

    }

}
