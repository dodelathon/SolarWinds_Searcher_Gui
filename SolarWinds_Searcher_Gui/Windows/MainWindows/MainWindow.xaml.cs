using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Management.Instrumentation;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;
using CustomExceptions;
using SolarWinds_Searcher_Gui.Windows.ErrorPopUps;
using SolarWinds_Searcher_Gui.Windows.MainWindows;
using SolarWinds_Searcher_Gui.Windows;

namespace SolarWinds_Searcher_Gui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        private SearcherThreadErrorPopUp searcherFail;
        private CustomParams customParams;
        private CountFailPopUp fail;
        private InvalidPathPopUp InvalidPath;
        private DircectoryDoesntExistOrIsEmptyPopUp empty;
        private Usage usage;
        private Confirmation prompt;
        private UnableToOpenPopup oFail;

        Runspace runspace;
        Pipeline pipe;

        private string user;
        private int col;
        private int ColsAcross;
        private int RowsAcross;
        private string Attribute;
        private int itemCount;
        private int row;
        private string path;

        private string[] Attributes;

        private bool Started = false;
        private bool Validated = false;
        private bool Refreshing = false;

        private ExcelInteraction excel;



        public MainWindow()
        {
            InitializeComponent();
            Visibility = Visibility.Hidden;
            prompt = new Confirmation(this);
            prompt.Visible = true;
        }

        public void Start()
        {
            user = Environment.UserName;
            EndExcel();
            Visibility = Visibility.Visible;
            ColsAcross = 10;
            RowsAcross = 10;
            Attribute = "Serial_Number";
            row = 2;
            col = 1;
            MinWidth = 440;
            MinHeight = 450;
            InitialConfig(0);
            Attributes = new string[]{ "Node Name", "IP Address", "IP Version", "DNS", "Machine Type", "Vendor", "Desciption", "Location", "Contact", "Status", "Software Image",
               "Software Version", "Asset_Environment", "Asset_Location", "Asset_Model", "Asset_State", "Cyber_Security_Classification", "Cybersecurity_Function",
               "Decom_Year", "DeviceType", "eFIApplication", "EOL_Date_HW", "EOL_Date_SW", "Hardware_Owner", "Holiday_Readiness", "Impact", "Imported_From_NCM", "InServiceDate",
               "Internet_Facing", "Legacy_Environment", "Local_Contact", "Management_Server", "Model_Category", "Network_Diagram", "Owner", "Physical_Host", "PONumber",
               "PurchaseDate", "PurchasePrice", "PurchasePrice_Maintenance", "QueueEmail", "Rack", "Rack_DataCenter", "Region", "Replacement_Cost", "Serial_Number",
               "SNOW_Assignment_Group", "SNOW_Configuration_Item", "SNOW_Product_Name", "Splunk_Index", "Splunk_Sourcetypes", "Term_End_Date", "Term_Start_Date", "Vendor"};
            Closing += Mainwindow_Closing;
            Task t = Task.Factory.StartNew(() => PopulateExcelCombo());
        }

        private void EndExcel()
        {
            RunspaceConfiguration configuration = RunspaceConfiguration.Create();
            runspace = RunspaceFactory.CreateRunspace(configuration);
            runspace.Open();
            RunspaceInvoke invoke = new RunspaceInvoke(runspace);
            pipe = runspace.CreatePipeline();
            pipe.Commands.AddScript("C:\\'Program Files'\\'SolarWinds Searcher'\\End-Excels.ps1");
            pipe.Invoke();

            pipe.Dispose();
            runspace.Close();
        }

        private void Mainwindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (excel != null)
            {
                excel.Close();
            }
        }

        /*private void SetPath()
        {
           
            const string name = "Path";
            string pathvar = Environment.GetEnvironmentVariable(name);
            var value = pathvar + "C:\\Program Files\\ChromeDriver";
            value += ";C:\\Program Files\\ChromeDriver\\WebDriver.dll";
            value += ";C:\\Program Files\\ChromeDriver\\WebDriver.Support.dll";
            if (!(pathvar.Equals(value)))
            {
                Console.WriteLine(value);

                var target = EnvironmentVariableTarget.Machine;
                Environment.SetEnvironmentVariable(name, value, target);
                Console.WriteLine(pathvar);
            }
        }*/

        private void InitialConfig(int mode)
        {
            if (mode == 0)
            {
                PathButton.Visibility = Visibility.Hidden;
                ExcelPathLabel.Visibility = Visibility.Hidden;
                ExcelPathBox.Visibility = Visibility.Hidden;
                ExcelPathBox.IsEnabled = false;
            }
            else if(mode == 1)
            {
                ExcelCombo.IsEnabled = false;
            }
            else if(mode == 2)
            {
                ExcelCombo.IsEnabled = true;
            }
            PathButton.IsEnabled = false;
            AttributeCombo.IsEnabled = false;
            Auto_SearchBox.IsEnabled = false;
            CustomBox.IsEnabled = false;
            SheetCombo.IsEnabled = false;
            ColCombo.IsEnabled = false;
            BeginBut.IsEnabled = false;
            CustomBox.IsChecked = false;
            Auto_SearchBox.IsChecked = false;
        }

        private void PopulateExcelCombo()
        {
            if (Directory.Exists("C:\\Users\\" + user + "\\Desktop\\Search_Repo") && (Directory.GetFiles("C:\\Users\\" + user + "\\Desktop\\Search_Repo\\", "*.xls*").Length != 0))
            {
                string[] res = Directory.GetFiles("C:\\Users\\" + user + "\\Desktop\\Search_Repo\\", "*.xls*");
                path = "C:\\Users\\" + user + "\\Desktop\\Search_Repo\\";

                Dispatcher.Invoke(() =>
                {
                    ExcelCombo.Items.Clear();
                    foreach (string o in res)
                    {
                        string[] temp = o.Split('\\');
                        ExcelCombo.Items.Add(temp[5]);
                    }
                    ExcelCombo.Items.MoveCurrentToFirst();
                });
            }
            else
            {
                Directory.CreateDirectory("C:\\Users\\" + user + "\\Desktop\\Search_Repo");
                empty = new DircectoryDoesntExistOrIsEmptyPopUp(this);
                empty.Visible = true;
            }
            Refreshing = false;

        }

        private void PopulateSheetCombo(string filename)
        {
            try
            {
                excel = new ExcelInteraction(filename);
                Collection<string> res = excel.GetSheets();

                Dispatcher.Invoke(() =>
                {
                    SheetCombo.Items.Clear();
                    foreach (string o in res)
                    {
                        SheetCombo.Items.Add(o.ToString());
                    }
                });
                if (Validated == true)
                {
                    Validated = false;
                    PathButton.IsEnabled = true;
                }
            }
            catch(Exception e)
            {
                Dispatcher.Invoke(() =>
                {
                    oFail = new UnableToOpenPopup(this);
                    oFail.Visible = true;
                    SheetCombo.IsEnabled = false;
                    ColCombo.IsEnabled = false;
                });
            }
            
        }

        private void PopulateColCombo()
        {
            Dispatcher.Invoke(() =>
            {
                ColCombo.Items.Clear();
                for (int i = 0; i < 100; i++)
                {
                    ColCombo.Items.Add(i);
                }
                ColCombo.Items.MoveCurrentToFirst();
                PopulateAttributeCombo();
            });

           

        }

        private void PopulateAttributeCombo()
        {
            Dispatcher.Invoke(() =>
            {
                AttributeCombo.Items.Clear();
                foreach (string i in Attributes)
                {
                    AttributeCombo.Items.Add(i);
                }
                AttributeCombo.Items.MoveCurrentToFirst();
            });
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Started == false)
            {
                Started = true;
                Dispatcher.Invoke(() =>
                {
                    InitialConfig(1);
                    Refresh.IsEnabled = false;
                });


                string SheetComboVal = SheetCombo.SelectedValue.ToString();
                string ExcelSheetVal = ExcelCombo.SelectedItem.ToString();
                if (excel == null)
                {
                    excel = new ExcelInteraction(ExcelSheetVal);
                }
                try
                {
                    excel.SetSheetName(SheetComboVal);
                    if (col == 11 || col == 12)
                    {
                        int[] arr = excel.FindStart(RowsAcross, ColsAcross, Attribute);
                        col = arr[1];
                        row = arr[0];
                    }
                    else
                    {
                        excel.SetCol(col);
                    }
                    itemCount = excel.GetCount(Attribute);
                    if (itemCount != -1)
                    {
                        Auto_SearchBox.IsEnabled = false;
                        Task t = Task.Factory.StartNew(() => DivvyExcel());
                    }
                    else
                    {
                        BeginBut.IsEnabled = true;
                        Refresh.IsEnabled = true;
                        fail = new CountFailPopUp(this);
                        fail.Activate();
                        fail.Visible = true;
                    }
                }
#pragma warning disable CS0168 // Variable is declared but never used
                catch (Exception ex)
#pragma warning restore CS0168 // Variable is declared but never used
                {
                    oFail = new UnableToOpenPopup(this);
                    oFail.Visible = true ;
                }
            }
        }


        private void Autosearch_Checked(object sender, RoutedEventArgs e)
        {
            col = 11;
            CustomBox.IsEnabled = true;
        }

        private void Autosearch_UnChecked(object sender, RoutedEventArgs e)
        {   if (!(ColCombo.Items.CurrentItem == null))
            {
                col = (int)ColCombo.SelectedItem;
            }
            else
            {
                col = 1;
            }
            CustomBox.IsChecked = false;
            CustomBox.IsEnabled = false;
        }

        private void ExcelCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ExcelCombo.SelectedItem != null)
            {
                string val = path + "\\" + ExcelCombo.SelectedItem.ToString();
                Console.WriteLine(val);
                Thread t = new Thread(() => PopulateSheetCombo(val));
                t.Start();
                SheetCombo.IsEnabled = true;
            }
        }

        private void SheetCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ColCombo.IsEnabled = true;
            PopulateColCombo();

        }

        private void FileInputSwitch_Checked(object sender, RoutedEventArgs e)
        {
            PathButton.Visibility = Visibility.Visible;
            PathButton.IsEnabled = true;
            ExcelPathBox.IsEnabled = true;
            ExcelPathBox.Visibility = Visibility.Visible;
            ExcelPathLabel.Visibility = Visibility.Visible;
            ExcelCombo.IsEnabled = false;
            ExcelComboLabel.Visibility = Visibility.Hidden;
            ExcelCombo.Visibility = Visibility.Hidden;
        }

        private void FileInputSwitch_UnChecked(object sender, RoutedEventArgs e)
        {
            PathButton.Visibility = Visibility.Hidden ;
            PathButton.IsEnabled = false;
            ExcelPathBox.IsEnabled = false;
            ExcelPathBox.Visibility = Visibility.Hidden;
            ExcelPathLabel.Visibility = Visibility.Hidden;
            ExcelCombo.IsEnabled = true;
            ExcelComboLabel.Visibility = Visibility.Visible;
            ExcelCombo.Visibility = Visibility.Visible;
            Task t = Task.Factory.StartNew(()=>PopulateExcelCombo());
            if(ExcelCombo.SelectedItem == null )
            {
                SheetCombo.IsEnabled = false;
            }
        }

        private void CustomBox_Checked(object sender, RoutedEventArgs e)
        {
            col = 12;
            customParams = new CustomParams(this, Attributes);
            customParams.Activate();
            customParams.Visibility = Visibility.Visible;    
        }

        private void CustomBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (Auto_SearchBox.IsChecked == true)
            {
                col = 11;
            }
            else
            {
                col = (int)ColCombo.SelectedValue;
            }
        }

        public void callback()
        {
            Object[] holder = customParams.GetAll();
            RowsAcross = (int)holder[0];
            ColsAcross = (int)holder[1];
            
        }

        private void ColCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            if (col != 101 && col != 102 && ColCombo.Items.IsEmpty != true)
            {
                col = int.Parse(ColCombo.SelectedItem.ToString());
                AttributeCombo.IsEnabled = true;
            }
            
        }

        private async void DivvyExcel()
        {
            //itemCount -= 4;
            double FirstQuarter = /*Math.Round*/((double)itemCount / 4);
            double SecondQuarter = /*Math.Round*/((double)(itemCount / 4) * 2);
            double ThirdQuarter = /*Math.Round*/((double)(itemCount / 4) * 3);
            
            SearcherThread searcher1;
            SearcherThread searcher2;
            SearcherThread searcher3;
            SearcherThread searcher4;

            bool found = false;
            int i = 0, AttributeIndex = 44;
            while(found == false)
            {
                Console.WriteLine(Attribute + " VS " + Attributes[i]);
                if(Attribute.Equals(Attributes[i]))
                {
                    found = true;
                    AttributeIndex = (i);
                }
                i++;
            }

            try
            {
                Console.WriteLine(col + " " + FirstQuarter + " " + itemCount);
                searcher1 = new SearcherThread(excel, col, row, (int)FirstQuarter, AttributeIndex);
                Console.WriteLine(col + " " + FirstQuarter + " " + SecondQuarter);
                searcher2 = new SearcherThread(excel, col, (int)FirstQuarter, (int)SecondQuarter, AttributeIndex);
                Console.WriteLine(col + " " + SecondQuarter + " " + ThirdQuarter);
                searcher3 = new SearcherThread(excel, col, (int)SecondQuarter, (int)ThirdQuarter, AttributeIndex);
                Console.WriteLine(col + " " + ThirdQuarter + " " + itemCount);
                searcher4 = new SearcherThread(excel, col, (int)ThirdQuarter, itemCount, AttributeIndex);

                Task runner1 = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        searcher1.SearchWrapper();
                    }
                    catch
                    {
                        Dispatcher.Invoke(() =>
                        {
                            searcherFail = new SearcherThreadErrorPopUp(this)
                            {
                                Visible = true
                            };

                        });
                    }
                });
                Task runner2 = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        searcher2.SearchWrapper();
                    }
                    catch
                    {
                        Dispatcher.Invoke(() =>
                        {
                            searcherFail = new SearcherThreadErrorPopUp(this)
                            {
                                Visible = true
                            };

                        });
                    }
                });
                Task runner3 = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        searcher3.SearchWrapper();
                    }
                    catch
                    {
                        Dispatcher.Invoke(() =>
                        {
                            searcherFail = new SearcherThreadErrorPopUp(this)
                            {
                                Visible = true
                            };                        
                        });
                    }
                });
                Task runner4 = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        searcher4.SearchWrapper();
                    }
                    catch
                    {
                        Dispatcher.Invoke(() =>
                        {
                            searcherFail = new SearcherThreadErrorPopUp(this)
                            {
                                Visible = true
                            };

                        });
                    }
                });

                await runner1;
                await runner2;
                await runner3;
                await runner4;
                Dispatcher.Invoke(() =>
                {
                    InitialConfig(2);
                    Refresh.IsEnabled = true;
                });

                excel.Save();
                excel.Show();
                PopulateExcelCombo();
                
            }
            catch (Exception e)
            {
                Dispatcher.Invoke(() =>
                {
                    BeginBut.IsEnabled = true;
                    Refresh.IsEnabled = true;
                    Started = false;
                    searcherFail = new SearcherThreadErrorPopUp(this);
                    searcherFail.Visible = true;

                });
            }
        }

        private void PathButton_Click(object sender, RoutedEventArgs e)
        {
            if (Validated == false)
            {
                Validated = true;
                PathButton.IsEnabled = false;
                string[] path = ExcelPathBox.Text.Split('\\');
                string dirPath = "";
                for (int i = 0; i < path.Length - 1; i++)
                {
                    if (i < path.Length - 2)
                    {
                        dirPath += path[i] + "\\";
                    }
                    else
                    {
                        dirPath += path[i];
                    }
                }
                Console.WriteLine(dirPath);
                if (Directory.Exists(dirPath))
                {
                    SheetCombo.IsEnabled = true;
                    PopulateSheetCombo(path[path.Length - 1]);
                }
                else
                {
                    PathButton.IsEnabled = true;
                    Validated = false;
                    SheetCombo.IsEnabled = false;
                    InvalidPath = new InvalidPathPopUp(this);
                    InvalidPath.Activate();
                    InvalidPath.Visible = true;
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (Refreshing == false)
            {
                Refreshing = true;
                Task t = Task.Factory.StartNew(() => PopulateExcelCombo());
                // t.Wait()
                BeginBut.IsEnabled = false;
                Auto_SearchBox.IsEnabled = true;
                Started = false;
            }
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            usage = new Usage(this);
            usage.Visible = true;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (AttributeCombo.SelectedValue != null)
            {
                Attribute = AttributeCombo.SelectedValue.ToString();
                Auto_SearchBox.IsEnabled = true;
                BeginBut.IsEnabled = true;
                Started = false;
            }
        }

        public void Dispose()
        {

        }
    }
}




/*

private static void PopulateExcelCombo()
{
    Collection<string> args = new Collection<string>();
    args.Add("-FileName");

    Collection<PSObject> res = RunScript("C:\\Users\\" + user + "\\Desktop\\SolarWinds_Searcher_Resources\\Fetch-Data.ps1", args);

    foreach (PSObject o in res)
    {
        Console.WriteLine(o);
    }
    PopulateSheetCombo(res[0].ToString());
}


private static void PopulateSheetCombo(String filename)
{
    Collection<string> args = new Collection<string>
            {
                "-SheetName," + "C:\\Users\\" + user + "\\Desktop\\Search_Repo\\" + filename ,
                "-Name"
            };

    Collection<PSObject> res = RunScript("C:\\Users\\" + user + "\\Desktop\\SolarWinds_Searcher_Resources\\Fetch-Data.ps1", args);

    foreach (PSObject o in res)
    {
        Console.WriteLine(o);
    }

}

    */