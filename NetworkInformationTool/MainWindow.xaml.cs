/* Title:           Network Information Tool
 * Date:            3-7-20
 * Author:          Terry Holmes
 * 
 * Description:     This is used to find network devices */

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
using System.Net.NetworkInformation;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace NetworkInformationTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();

        IPAddressesDataSet TheIPAddressesDataSet = new IPAddressesDataSet();


        public MainWindow()
        {
            InitializeComponent();
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void expFindDevices_Expanded(object sender, RoutedEventArgs e)
        {
            //setting up local variables
            int intTimeOut = 1;
            int intCounter;
            Ping pngNewPing = new Ping();
            PingReply repPing;
            string strAddress;
            string strSuccess = "";
            IPHostEntry HostEntry;
            string strHostName = "";

            expFindDevices.IsExpanded = false;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                TheIPAddressesDataSet.ipaddress.Rows.Clear();

                if(txtEnterIPAddress.Text == "")
                {
                    TheMessagesClass.ErrorMessage("IP Address Range Not Found");
                    return;
                }

                for(intCounter = 0; intCounter <= 255; intCounter++)
                {
                    strAddress = txtEnterIPAddress.Text + Convert.ToString(intCounter);

                    repPing = pngNewPing.Send(strAddress, intTimeOut);

                    if(repPing.Status == IPStatus.Success)
                    {
                        try
                        {
                            strSuccess = "Sucess";

                            HostEntry = Dns.GetHostEntry(strAddress);

                            strHostName = HostEntry.HostName;
                        }
                        catch (Exception)
                        {
                            strHostName = "NOT FOUND";
                        }

                        IPAddressesDataSet.ipaddressRow NewAddressRow = TheIPAddressesDataSet.ipaddress.NewipaddressRow();

                        NewAddressRow.Address = strAddress;
                        NewAddressRow.Success = strSuccess;
                        NewAddressRow.HostName = strHostName;

                        TheIPAddressesDataSet.ipaddress.Rows.Add(NewAddressRow);

                        dgrResults.ItemsSource = TheIPAddressesDataSet.ipaddress;
                    }
                    else
                    {
                        strSuccess = "Failure";
                    }
                    
                }
            }
            catch (Exception Ex)
            {
                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();
            
        }

        private void expExportToExcel_Expanded(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                expExportToExcel.IsExpanded = false;

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheIPAddressesDataSet.ipaddress.Rows.Count;
                intColumnNumberOfRecords = TheIPAddressesDataSet.ipaddress.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheIPAddressesDataSet.ipaddress.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheIPAddressesDataSet.ipaddress.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
