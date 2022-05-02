using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace AdrianCanas
{
    public partial class Main : Form
    {
        public static string folderParaMonitorear = "";
        public FileSystemWatcher watcher = new FileSystemWatcher();
        public static ListBox ListFiles = null;
        public static Microsoft.Office.Interop.Excel.Application oXL;
        public static  Microsoft.Office.Interop.Excel.Workbook oWB;
        public static Microsoft.Office.Interop.Excel.Worksheet oSheet;
        public static Microsoft.Office.Interop.Excel.Workbook oNewWB;
        public static Microsoft.Office.Interop.Excel.Worksheet oNewSheet;

        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                folderParaMonitorear = folderBrowserDialog1.SelectedPath;
                //folderParaMonitorear = @"C:\\Code\\Folders";
                watcher.Path = folderParaMonitorear;
                label1.Text = folderParaMonitorear;
                

                watcher.NotifyFilter = NotifyFilters.Attributes |
                                        NotifyFilters.CreationTime |
                                        NotifyFilters.DirectoryName |
                                        NotifyFilters.FileName |
                                        NotifyFilters.LastAccess |
                                        NotifyFilters.LastWrite |
                                        NotifyFilters.Security |
                                        NotifyFilters.Size;
                // Watch all files.  
                watcher.Filter = "*.*";
                // Add event handlers.  
                //watcher.Changed += new FileSystemEventHandler(OnChanged);
                watcher.Created += new FileSystemEventHandler(OnChanged);
                //watcher.Deleted += new FileSystemEventHandler(OnChanged);
                //watcher.Renamed += new RenamedEventHandler(OnRenamed);
                //Start monitoring.  
                watcher.EnableRaisingEvents = true;

                ListFiles = listBox1;

                FillItemsInListBox(folderParaMonitorear);

                Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(folderParaMonitorear), "Processed"));
                Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(folderParaMonitorear), "Not Applicable"));

                //oXL.Application.DisplayAlerts = false;
            }
            // Define the event handlers.  

        }
        public static void OnChanged(object source, FileSystemEventArgs e)
        {
            string DestArch = "";

            // Specify what is done when a file is changed.  
            Console.WriteLine("{0}, with path {1} has been {2}", e.Name, e.FullPath, e.ChangeType);
            if (e.ChangeType == WatcherChangeTypes.Created)
            {


                DestArch = Directory.GetParent(Directory.GetParent(e.FullPath).ToString()).ToString();



                if (Path.GetExtension(e.FullPath).Equals(".xls") || Path.GetExtension(e.FullPath).Equals("..xlsm") || Path.GetExtension(e.FullPath).Equals(".xlsx"))
                {
                    DestArch = Path.Combine(DestArch, "Processed", Path.GetFileName(e.FullPath));
                    if (!File.Exists(DestArch))
                    {
                        while (!FileIsReady(e.FullPath)){}
                        File.Move(e.FullPath, DestArch);
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = false;


                        string MasterWorkbookPath = "";
                        MasterWorkbookPath= Path.Combine( Directory.GetParent(DestArch).ToString(),"MasterWorkBook.xlsx" );
                        if (File.Exists(MasterWorkbookPath))
                        {
                            oWB = oXL.Workbooks.Open(MasterWorkbookPath);
                        }
                        else
                        {
                            oWB = oXL.Workbooks.Add();
                            oWB.SaveCopyAs(MasterWorkbookPath);
                            
                        }

                        oNewWB=oXL.Workbooks.Open(DestArch);

                        int totalSheets = oWB.Sheets.Count;

                        //oWB.Worksheets.Add("NewSheet1");

                        oWB.Worksheets.Add(
    System.Reflection.Missing.Value,
    oWB.Worksheets[oWB.Worksheets.Count],
    1,
    System.Reflection.Missing.Value);
                        oWB.Save();

                        //foreach (Microsoft.Office.Interop.Excel.Worksheet sht in oNewWB.Worksheets)
                        //{
                        //    //oWB.Worksheets.Add(sht);
                        //    //oWB.Sheets.Add(sht,After:     );
                        //    sht.Move(oWB.Sheets[totalSheets]);
                        //    totalSheets ++;

                        //    oWB.Save();

                        //}

                        oNewWB.Close();
                        oWB.Close();
                        oXL.Quit();



                    }
                    else
                    {
                        while (!FileIsReady(e.FullPath)) { }
                        File.Delete(e.FullPath);
                    }
                }
                else
                {
                    DestArch = Path.Combine(DestArch, "Not Applicable", Path.GetFileName(e.FullPath));
                    if (!File.Exists(DestArch))
                    {
                        while (!FileIsReady(e.FullPath)) { }
                        File.Move(e.FullPath, DestArch);
                    }
                    else
                    {
                        while (!FileIsReady(e.FullPath)) { }
                        File.Delete(e.FullPath);
                    }
                }
            }
        }
    public static void OnRenamed(object source, RenamedEventArgs e)
    {
        // Specify what is done when a file is renamed.  
        Console.WriteLine(" {0} renamed to {1}", e.OldFullPath, e.FullPath);
    }
    public static void FillItemsInListBox(string pathFolder)
    {
        foreach (string FileInDirectory in Directory.GetFiles(folderParaMonitorear))
        {
            ListFiles.Items.Add(FileInDirectory);
        }
    }
        public static bool FileIsReady(string path)
        {
            //One exception per file rather than several like in the polling
            try
            {
                //if we can' open the file, it's still copying
                using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return true;
                }
            }
            catch (IOException)
            {
                return false;
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            oXL = new Microsoft.Office.Interop.Excel.Application();

            

            string masterWorkbookPath = "C:\\Code\\Master.xlsx";

            if (File.Exists(masterWorkbookPath))
            {
                oWB = oXL.Workbooks.Open(masterWorkbookPath);
            }
            else
            {
                oWB = oXL.Workbooks.Add();
                
                oWB.SaveCopyAs(masterWorkbookPath);

            }
            oNewWB = oXL.Workbooks.Open("C:\\Code\\Master2.xlsx");

            oNewSheet = oNewWB.Worksheets[0];

            Microsoft.Office.Interop.Excel.Worksheet newWorksheet;

            //oWB.Worksheets.Add(oNewSheet);
            oWB.Sheets.Add(oNewSheet);

            oWB.Save();
            oWB.Close();
            oXL.Quit();

            oNewWB.Save();
            oNewWB.Close();




        }
    }

}
