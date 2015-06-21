using System;
using System.ServiceProcess;
using System.Diagnostics;

namespace ModbusRTUService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            try
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
                { 
                    new ModbusRTUService() 
                };
                ServiceBase.Run(ServicesToRun);
            }
            catch(Exception ex)
            {
                EventLog eventLog = new EventLog();
                if (!EventLog.SourceExists("ModbusRTUService"))
                {
                    EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService");
                }
                eventLog.Source = "ModbusRTUService";
                eventLog.WriteEntry(String.Format("Exception: {0} \n\nStack: {1}", ex.Message + " : " + ex.ToString(), ex.StackTrace) , EventLogEntryType.Error );
            }
        }
    }
}
