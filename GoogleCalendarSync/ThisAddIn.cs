using System;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace GoogleCalendarSync
{
    public partial class ThisAddIn
    {
        private GoogleCalendarAPI api;
        private Outlook.Inspectors inspectors;
        private Outlook.AppointmentItem appointment;
        private Outlook.Folder folder;
        DateTime _start;
        DateTime _end;
        string _subject;
        string _id;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            api = new GoogleCalendarAPI();

            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            folder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
        }

        private Outlook.AppointmentItem GetAppointment(string id)
        {
            // fix this
            string filter = "[GlobalAppointmentID] = '" + id +"'";
            Debug.WriteLine(filter);

            if (folder.Items.Find(filter) != null)
            {
                return folder.Items.Find(filter) as Outlook.AppointmentItem;
            }
            else return null;
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            var item = Inspector.CurrentItem as Outlook.AppointmentItem;
            if (item != null)
            {
                appointment = item;
                appointment.Write += Appointment_Write;
                appointment.BeforeDelete += Appointment_BeforeDelete;
            }
            
        }

        private void Appointment_BeforeDelete(object Item, ref bool Cancel)
        {
            _id = appointment.GlobalAppointmentID.ToLower();
            if (!api.DeleteAppointment(_id))
            {
                Cancel = true;
            }
        }

        private void Appointment_Write(ref bool Cancel)
        {
            _start = appointment.Start;
            _end = appointment.End;
            _subject = appointment.Subject;
            _id = appointment.GlobalAppointmentID;

            if (appointment.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled)
            {
                Cancel = true;
            }
            if (GetAppointment(appointment.GlobalAppointmentID) != null)
            {
                _id = _id.ToLower();
                api.UpdateAppointment(_id, _start, _end, _subject);
            }
            else
            {
                _id = _id.ToLower();
                api.NewAppointment(_start, _end, _subject, _id);
            }
            
            Marshal.ReleaseComObject(appointment);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
