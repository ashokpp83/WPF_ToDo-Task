﻿using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Microsoft.Office.Interop;
using System.IO;
using System.ComponentModel;

namespace WPF_ToDoList
{
    public partial class TabViewModel
    {
        public string Name { get; set; }
        public ObservableCollection<TabItemViewModel> Collection { get; set; }
    }

    public partial class TabItemViewModel
    {
        public string Title { get; set; }
        public SolidColorBrush background { get; set; }
        //public string Detail { get; set; }
    }

    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private string ExportImportFileName = "ToDoList_ExportImport.xls";
        public ObservableCollection<TabViewModel> tabViewModels { get; set; }

        //List<string> list_content = new List<string>();
        public MainWindow()
        {
            InitializeComponent();

            tabViewModels = new ObservableCollection<TabViewModel>();

            string path = System.AppDomain.CurrentDomain.BaseDirectory + ExportImportFileName;
            if (File.Exists(path))
            {
                this.Cursor = Cursors.Wait;

                ImportFromExcel(path);

                this.Cursor = Cursors.Arrow;
            }
        }
        
        private void btnAddTab_Click(object sender, RoutedEventArgs e)
        {

            string tabname = string.Empty;
            TabProperty dlg = new TabProperty();

            // get existing header text
            //dlg.txtTitle.Text = tab.Header.ToString();

            if (dlg.ShowDialog() == true)
            {
                // change header text
                tabname = dlg.txtTitle.Text.Trim();
            }

            tabViewModels.Add(new TabViewModel { Name = tabname, Collection = new ObservableCollection<TabItemViewModel> { new TabItemViewModel { Title = "", background = Brushes.Transparent } } }); 

            DataContext = tabViewModels;

            tabControlName.SelectedIndex = tabControlName.Items.Count -1;
        }

        private void btnDeleteTab_Click(object sender, RoutedEventArgs e)
        {
            TabViewModel tvm = tabControlName.SelectedItem as TabViewModel;

            tabViewModels.Remove(tvm);

            DataContext = tabViewModels;

        }

        private void btnAddTask_Click(object sender, RoutedEventArgs e)
        {
            string task = txt_Task.Text;

            if (string.IsNullOrEmpty(task))
            {
                MessageBox.Show("Please enter Task to Add", "", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            TabViewModel tvm = tabControlName.SelectedItem as TabViewModel ;
            
            LoadData(tvm, task);

        }

        private void LoadData(TabViewModel tvm, string task)
        {
            int index = tabControlName.SelectedIndex;

            List<TabItemViewModel> list_content = new List<TabItemViewModel>();
            foreach (TabItemViewModel item in tvm.Collection)
            {
                if (item.Title.Length > 0)
                {
                    TabItemViewModel tvi = new TabItemViewModel();
                    tvi.Title = item.Title;
                    tvi.background = item.background;

                    list_content.Add(tvi);
                }
            }

            //new task
            TabItemViewModel tvi2 = new TabItemViewModel();
            tvi2.Title = task;
            tvi2.background = Brushes.Transparent;

            //new item
            list_content.Add(tvi2);

            tabViewModels.Remove(tvm);

            ObservableCollection<TabItemViewModel> obs = new ObservableCollection<TabItemViewModel>();

            foreach (TabItemViewModel cont in list_content)
            {   
                obs.Add(cont);
            }

            tabViewModels.Insert(index, new TabViewModel { Name = tvm.Name, Collection = obs });

            //tabViewModels.
            DataContext = tabViewModels;

            tabControlName.SelectedIndex = index;

        }

        private void btnDeleteTask_Click(object sender, RoutedEventArgs e)
        {
            //selected tab
            TabViewModel tvm = tabControlName.SelectedItem as TabViewModel;

            List<TabItemViewModel> list_content = new List<TabItemViewModel>();
            foreach (TabItemViewModel item in tvm.Collection)
            {
                //remove the selected item
                if (item.Title.Length > 0 && item.Title != lb_selectedItem)
                {
                    TabItemViewModel tiv = new TabItemViewModel();
                    tiv.Title = item.Title;
                    tiv.background = item.background;
                    
                    list_content.Add(tiv);
                }
            }

            int index = tabControlName.SelectedIndex;

            tabViewModels.Remove(tvm);

            ObservableCollection<TabItemViewModel> obs = new ObservableCollection<TabItemViewModel>();
            foreach (TabItemViewModel cont in list_content)
            {
                obs.Add(cont);
            }

            tabViewModels.Insert(index, new TabViewModel { Name = tvm.Name, Collection = obs });
            
            //tabViewModels.
            DataContext = tabViewModels;

            //tabControlName.SelectedItem = tvm;
            tabControlName.SelectedIndex = index;

        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "xlsx Files (*.xlsx)|*.xls|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == true)
            {
                this.Cursor = Cursors.Wait;

                string filename = openFileDialog1.FileName;
                ImportFromExcel(filename);

                this.Cursor = Cursors.Arrow;
            }

        }
       private void ImportFromExcel(string path)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                tabViewModels.Clear();

                List<ExcelData> datalist = new List<ExcelData>();

                //skip row 1 for header
                for (int i = 2; i <= xlRange.Rows.Count; i++)
                {
                    ExcelData data = new ExcelData();

                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range range = (xlWorksheet.Cells[i, j] as Microsoft.Office.Interop.Excel.Range);
                        string cellValue = range.Value.ToString();

                        if (j == 1)
                            data.Tab = cellValue;
                        if (j == 2)
                            data.Task = cellValue;
                        if (j == 3)
                        {
                            if (cellValue == "Completed")
                                data.Status = cellValue; //tivm.background = Brushes.Green;
                            else
                                data.Status = cellValue; //tivm.background = Brushes.Transparent;
                        }

                    }

                    datalist.Add(data);
                }

                List<string> uniquetabs = new List<string>();
                foreach (ExcelData d in datalist)
                {
                    if (!uniquetabs.Contains(d.Tab))
                        uniquetabs.Add(d.Tab);
                }

                foreach (string tab in uniquetabs)
                {
                    TabViewModel tvm = new TabViewModel();
                    ObservableCollection<TabItemViewModel> obs = new ObservableCollection<TabItemViewModel>();

                    foreach (ExcelData d in datalist)
                    {
                        if (d.Tab == tab)
                        {
                            TabItemViewModel tiv = new TabItemViewModel();
                            tiv.Title = d.Task;
                            if (d.Status == "Completed")
                                tiv.background = Brushes.Green;
                            else
                                tiv.background = Brushes.Transparent;

                            obs.Add(tiv);
                        }
                    }

                    tvm.Name = tab;
                    tvm.Collection = obs;

                    tabViewModels.Add(tvm);
                }

                xlWorkbook.Close();
                xlApp.Quit();

                DataContext = tabViewModels;

                tabControlName.SelectedIndex = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                //GC.WaitForFullGCComplete();
                GC.WaitForPendingFinalizers();
            }
        }



        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.InitialDirectory = @"C:\";
            saveDlg.Filter = "Excel files (*.xls)|*.xls";
            saveDlg.FilterIndex = 0;
            saveDlg.RestoreDirectory = true;
            saveDlg.Title = "Export Excel File To";
            if (saveDlg.ShowDialog() == true)
            {
                this.Cursor = Cursors.Wait;

                string path = saveDlg.FileName;
                ExportToExcel(path);

                this.Cursor = Cursors.Arrow;

                MessageBox.Show("Succssfully Exported the data to " + path, "Export File", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }

        private void ExportToExcel(string path, bool DisplayAlert = true)
        { 
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                if (!DisplayAlert)
                    xlApp.DisplayAlerts = false;

                xlWorkSheet.Cells[1, 1] = "Category";
                xlWorkSheet.Cells[1, 2] = "Task";
                xlWorkSheet.Cells[1, 3] = "Status";

                int row = 2;
                foreach (TabViewModel tvm in tabViewModels)
                {   
                    //xlWorkSheet.Name = tvm.Name;
                    foreach (TabItemViewModel tivm in tvm.Collection)
                    {
                        xlWorkSheet.Cells[row, 1] = tvm.Name;

                        xlWorkSheet.Cells[row, 2] = tivm.Title;
                        if (tivm.background == Brushes.Transparent)
                            xlWorkSheet.Cells[row, 3] = "InProgress";
                        else
                            xlWorkSheet.Cells[row, 3] = "Completed";

                        row++;
                    }                    
                }

                //if(!File.Exists(path) )
                //    File.Create(path);

                xlWorkBook.SaveAs(path);
                xlWorkBook.Saved = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception " + ex.Message);                
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                //GC.WaitForFullGCComplete();
                GC.WaitForPendingFinalizers();
            }
        }

        private void btnToggleTask_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(lb_selectedItem))
                MessageBox.Show("Please selet the task to Mark as Complete/Not Complete", "Toggle Task", MessageBoxButton.OK, MessageBoxImage.Warning);   

            TabViewModel tvm = tabControlName.SelectedItem as TabViewModel;

            List<TabItemViewModel> list_content = new List<TabItemViewModel>();
            foreach (TabItemViewModel item in tvm.Collection)
            {
                //TabItemViewModel tiv = new TabItemViewModel();
                //remove the selected item
                if (item.Title == lb_selectedItem)
                {  
                    //tiv.Title = item.Title;
                    if (item.background == Brushes.Green)
                        item.background = Brushes.Transparent;
                    else
                        item.background = Brushes.Green;
                }
                list_content.Add(item);
            }

            int index = tabControlName.SelectedIndex;

            tabViewModels.Remove(tvm);

            ObservableCollection<TabItemViewModel> obs = new ObservableCollection<TabItemViewModel>();
            foreach (TabItemViewModel cont in list_content)
            {
                obs.Add(cont);
            }

            tabViewModels.Insert(index, new TabViewModel { Name = tvm.Name, Collection = obs });

            //tabViewModels.
            DataContext = tabViewModels;

            //tabControlName.SelectedItem = tvm;
            tabControlName.SelectedIndex = index;            
        }

        private string lb_selectedItem = string.Empty;
        private System.Windows.Controls.ListBox lstBox;
        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                lstBox = sender as System.Windows.Controls.ListBox;

                if (lstBox.SelectedItem != null)
                    lb_selectedItem = (lstBox.SelectedItem as TabItemViewModel).Title;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception " + ex.Message);
            }

        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            string path  = System.AppDomain.CurrentDomain.BaseDirectory+ ExportImportFileName;

            this.Cursor = Cursors.Wait;

            //to load when the app
            ExportToExcel(path, false);

            this.Cursor = Cursors.Arrow;
        }
    }

    public class ExcelData
    {
        public string Tab { get; set; }
        public string Task{ get; set; }
        public string Status{ get; set; }
    }
}
