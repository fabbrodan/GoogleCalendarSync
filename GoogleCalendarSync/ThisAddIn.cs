using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace GoogleCalendarSync
{
    public partial class ThisAddIn
    {
        public GoogleCalendarAPI api;
        public Outlook.Inspectors inspectors;
        public Outlook.AppointmentItem appointment;
        DateTime _start;
        DateTime _end;
        string subject;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            api = new GoogleCalendarAPI();
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            var item = Inspector.CurrentItem as Outlook.AppointmentItem;
            if (item != null)
            {
                appointment = item;
                appointment.AfterWrite += Appointment_AfterWrite;
                appointment.PropertyChange += Appointment_PropertyChange;
            }
            
        }

        private void Appointment_PropertyChange(string Name)
        {
            if (Name == "StartInStartTimeZone")
            {
                _start = appointment.Start;
            }
            if (Name == "EndInEndTimeZone")
            {
                _end = appointment.End;
            }
            if (Name == "Subject")
            {
                subject = appointment.Subject;
            }
        }

        private void Appointment_AfterWrite()
        {
            api.NewAppointment(_start, _end, subject);
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
