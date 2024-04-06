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
using DateSearchDLL;
using NewEventLogDLL;
using NewEmployeeDLL;

namespace AutoEmailFileReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();

        //setting up the data sets
        FindActiveServerLogSearchTermsDataSet TheFindActiveServerLogSearchTermsDataSet = new FindActiveServerLogSearchTermsDataSet();
        FindServerEventLogForReportsByItemDataSet TheFindServerEventLogForReportsByItemDataSet = new FindServerEventLogForReportsByItemDataSet();
        FindServerEventLogForReportsByUserDataSet TheFindServerEventLogForReportsByUserDataSet = new FindServerEventLogForReportsByUserDataSet();
        FindRandomEmployeeDataSet TheFindRandomEmployeeDataSet = new FindRandomEmployeeDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            string strSearchTerm;
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            string strLastName;

            try
            {
                TheFindActiveServerLogSearchTermsDataSet = TheEventLogClass.FindActiveServerLogSearchTerms();

                intNumberOfRecords = TheFindActiveServerLogSearchTermsDataSet.FindActiveServerLogSearchTerms.Rows.Count;

                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDateSearchClass.AddingDays(datStartDate, 1);
                datStartDate = TheDateSearchClass.SubtractingDays(datStartDate, 1);

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strSearchTerm = TheFindActiveServerLogSearchTermsDataSet.FindActiveServerLogSearchTerms[intCounter].SearchTerm;

                        TheFindServerEventLogForReportsByUserDataSet = TheEventLogClass.FindServerEventLogForReportsByUser(strSearchTerm, datStartDate, datEndDate);

                        intRecordsReturned = TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser.Rows.Count;

                        if(intRecordsReturned < 1)
                        {
                            TheFindServerEventLogForReportsByItemDataSet = TheEventLogClass.FindServerEventLogForReportsByItem(strSearchTerm, datStartDate, datEndDate);

                            intRecordsReturned = TheFindServerEventLogForReportsByItemDataSet.FindServerEventLogForReportsByItem.Rows.Count;

                            if(intRecordsReturned > 0)
                            {
                                SendItemReport();
                            }
                        }
                        else if(intRecordsReturned > 0)
                        {
                            SendUserReport();
                        }
                    }
                }

                TheFindRandomEmployeeDataSet = TheEmployeeClass.FindRandomEmployees();

                intNumberOfRecords = TheFindRandomEmployeeDataSet.FindRandomEmployees.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strLastName = TheFindRandomEmployeeDataSet.FindRandomEmployees[intCounter].LastName;

                        TheFindServerEventLogForReportsByUserDataSet = TheEventLogClass.FindServerEventLogForReportsByUser(strLastName, datStartDate, datEndDate);

                        intRecordsReturned = TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser.Rows.Count;

                        if(intRecordsReturned > 0)
                        {
                            SendUserReport();
                        }
                    }
                }

                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Auto Email File Report // Main Window // Window Loaded Method " + ex.Message);

                TheMessagesClass.ErrorMessage(ex.ToString());
            }
        }
        private void SendUserReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "itadmin@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            DateTime datPayDate = DateTime.Now;
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser.Rows.Count;

                strHeader = "Server File Access Report Prepared on " + Convert.ToString(datPayDate);

                strMessage = "<h1>Server File Access Report Prepared on " + Convert.ToString(datPayDate) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Transaction Date</b></td>";
                strMessage += "<td><b>Logon Name</b></td>";
                strMessage += "<td><b>Item Accessed</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + Convert.ToString(TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser[intCounter].TransactionDate) + "</td>";
                        strMessage += "<td>" + TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser[intCounter].LogonName + "</td>";
                        strMessage += "<td>" + TheFindServerEventLogForReportsByUserDataSet.FindServerEventLogForReportsByUser[intCounter].ItemAccessed + "</td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = !(TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage));

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Auto Email File Report // Main Window // Send User Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void SendItemReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "itadmin@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            DateTime datPayDate = DateTime.Now;
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheFindServerEventLogForReportsByItemDataSet.FindServerEventLogForReportsByItem.Rows.Count;

                strHeader = "Server File Access Report Prepared on " + Convert.ToString(datPayDate);

                strMessage = "<h1>Server File Access Report Prepared on " + Convert.ToString(datPayDate) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Transaction Date</b></td>";
                strMessage += "<td><b>Logon Name</b></td>";
                strMessage += "<td><b>Item Accessed</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + Convert.ToString(TheFindServerEventLogForReportsByItemDataSet.FindServerEventLogForReportsByItem[intCounter].TransactionDate) + "</td>";
                        strMessage += "<td>" + TheFindServerEventLogForReportsByItemDataSet.FindServerEventLogForReportsByItem[intCounter].LogonName + "</td>";
                        strMessage += "<td>" + TheFindServerEventLogForReportsByItemDataSet.FindServerEventLogForReportsByItem[intCounter].ItemAccessed + "</td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = !(TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage));

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Auto Email File Report // Main Window // Send Item Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

    }
}
