/* Title:           Send Email Class
 * Date:            7-12-17
 * Author:          Terry Holmes */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using NewEventLogDLL;
using System.Configuration;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.IO;
using SafetyDLL;


namespace AutoEmailFileReport
{
    class SendEmailClass
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        SafetyClass TheSafetyClass = new SafetyClass();

        VehicleInspectionEmailListDataSet TheVehicleInspectionEmailListDataSet;


        public void TrailerEmailMessage(string strVehicleNumber, string strVehicleProblem)
        {
            bool blnEmailSent;
            string strVehicleFailed;
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress;
            string strHeader = "Trailer Inspection Failure - Do Not Reply";

            strVehicleFailed = "<h1>" + strHeader + "</h1>" + "<h3>" + "Trailer Number " + strVehicleNumber + "</h3>" + "<p> Has Failed The Trailer Inspection and Can Not Be Used " + strVehicleProblem + "</p>";

            TheVehicleInspectionEmailListDataSet = TheSafetyClass.GetVehicleInspectionEmailListInfo();

            intNumberOfRecords = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist.Rows.Count - 1;

            for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                strEmailAddress = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist[intCounter].EmailAddress;

                blnEmailSent = SendEmail(strEmailAddress, strHeader, strVehicleFailed);
            }
        }
        public void EmailMessage(string strVehicleNumber, string strVehicleProblem)
        {
            bool blnEmailSent;
            string strVehicleFailed;
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress;
            string strHeader = "Vehicle Inspection Failure - Do Not Reply";

            strVehicleFailed = "<h1>" + strHeader + "</h1>" + "<h3>" + "Vehicle Number " + strVehicleNumber + "</h3>"  + "<p> Has Failed The Vehicle Inspection and Can Not Be Driven " + strVehicleProblem + "</p>";

            TheVehicleInspectionEmailListDataSet = TheSafetyClass.GetVehicleInspectionEmailListInfo();

            intNumberOfRecords = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist.Rows.Count - 1;

            for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                strEmailAddress = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist[intCounter].EmailAddress;

                blnEmailSent = SendEmail(strEmailAddress, strHeader, strVehicleFailed);
            }
        }
         
        public bool SendEmail(string mailTo, string subject, string message)
        {
            

            try
            {

                MailMessage mailMessage = new MailMessage("bluejayerpreports@bluejaycommunications.com", mailTo, subject, message);
                mailMessage.IsBodyHtml = true;
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.SubjectEncoding = Encoding.UTF8;

                SmtpClient smtpClient = new SmtpClient("192.168.0.240", 25);
                smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = false;
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpClient.Send(mailMessage);
                return true;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Send Email Class // Send Mail " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                return false;
            }


        }
        public void VehicleReports(string strHeader, String strVehicleReport)
        {
            bool blnEmailSent;
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress;

            TheVehicleInspectionEmailListDataSet = TheSafetyClass.GetVehicleInspectionEmailListInfo();

            intNumberOfRecords = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist.Rows.Count - 1;

            for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                strEmailAddress = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist[intCounter].EmailAddress;

                blnEmailSent = SendEmail(strEmailAddress, strHeader, strVehicleReport);
            }
        }
    }
}
