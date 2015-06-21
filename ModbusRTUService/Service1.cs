using System;
using System.ComponentModel;
using System.ServiceProcess;
using System.Timers;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;

namespace ModbusRTUService
{
    public partial class ModbusRTUService : ServiceBase
    {

        public EventLog eventLog = new EventLog();      // Переменная для записи в журнал событий
        private System.Timers.Timer timerSrv;           // Таймер периодичности опроса
        List<string> unitA3Files = new List<string>();  // Список файлов аналоговых сигналов для Slave 3
        List<string> unitA4Files = new List<string>();  // Список файлов аналоговых сигналов для Slave 4
        List<string> unitB3Files = new List<string>();  // Список файлов дискретных сигналов для Slave 3
        List<string> unitB4Files = new List<string>();  // Список файлов дискретных сигналов для Slave 4
        private ushort[] AWAUS3;                        // Переменная для аналоговых значений Slave 3 из файла AWAUS_UNIT3 
        private ushort[] BWAUS3;                        // Переменная для дискретных значений Slave 3 из файла BWAUS_UNIT3
        private ushort[] AWAUS4;                        // Переменная для аналоговых значений Slave 4 из файла AWAUS_UNIT4
        private ushort[] BWAUS4;                        // Переменная для дискретных значений Slave 4 из файла BWAUS_UNIT3
        // private ushort[] pto_1;                      // Переменная для аналоговых значений pto_1
        private IModbusService mbSlave;                 // Класс для трансляции данных в Modbus
        private IFileParse fileParse;                   // Класс для обработки файлов и записи их переменные
        private Thread threadSlave;                     // Поток, в котором будет работать Modbus

        // Инициализация службы
        public ModbusRTUService()
        {
            #region Инициализация компонентов

            InitializeComponent();
            
            //Добавляем список файлов для аналоговых сигналов Slave 3
            unitA3Files.Add(@"C:\unit_3_4\AWAUS_UNIT3");
            unitA3Files.Add(@"C:\unit_3_4\AWAUS_UNIT3_1");

            //Добавляем список файлов для аналоговых сигналов Slave 4
            unitA4Files.Add(@"C:\unit_3_4\AWAUS_UNIT4");

            //Добавляем список файлов для аналоговых сигналов Slave 3
            unitB3Files.Add(@"C:\unit_3_4\BWAUS_UNIT3");

            //Добавляем список файлов для аналоговых сигналов Slave 4
            unitB4Files.Add(@"C:\unit_3_4\BWAUS_UNIT4");

            //Отключаем автоматическую запись в журнал
            AutoLog = false;

            // Создаем журнал событий и записываем в него
            if (!EventLog.SourceExists("ModbusRTUService")) //Если журнал с таким названием не существует
            {
                EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService"); // Создаем журнал
            }
            eventLog.Source = "ModbusRTUService"; //Помечаем, что будем писать в этот журнал

            fileParse = new FileParse();    //Инициализация класса для обработки файлов
            mbSlave = new ModbusService();  //Инициализация класса трансляции данных по Modbus

            #endregion
        }

        // Запуск службы
        protected override void OnStart(string[] args)
        {
            #region Запись в журнал

            eventLog.WriteEntry("Служба запущена");

            #endregion

            #region Инициализация таймера
            //Инициализация таймера
            timerSrv = new System.Timers.Timer(10000);
            //Задание интервала опроса
            timerSrv.Interval = 60000;
            //Включение таймера
            timerSrv.Enabled = true;
            //Добавление обработчика на таймер
            timerSrv.Elapsed += new ElapsedEventHandler(ReadAndModbus);
            //Автоматический взвод таймера 
            timerSrv.AutoReset = true;
            //Старт таймера
            timerSrv.Start();
   
            #endregion
        }

        protected override void OnStop()
        {
            
            #region Запись в журнал

            eventLog.WriteEntry("Служба остановлена");

            #endregion
        }
        // Обработчик таймера
        private void ReadAndModbus(object sender, ElapsedEventArgs e)
        {
            #region Обработка файлов и запись их в переменнные

            AWAUS3 = fileParse.AWAUSParse(unitA3Files);                 //Обработка файла аналоговых значений для Slave 3
            BWAUS3 = fileParse.BWAUSParse(unitB3Files);                 //Обработка файла дискретных значений для Slave 3
            AWAUS4 = fileParse.AWAUSParse(unitA4Files);                 //Обработка файла аналоговых значений для Slave 4
            BWAUS4 = fileParse.BWAUSParse(unitB4Files);                 //Обработка файла дискретных значений для Slave 4
            // pto_1 = fileParse.ptoParse(@"C:\unit_3_4\pto_1.txt");    //Обработка файла дискретных значений для pto

            #endregion

            #region Перезапуск потока Modbus

            StopSlaveThread();      //Остановска службы, если она запущена
            InitThreads();          //Инициализация данных потока
            threadSlave.Start();    //Запуск потока

            #endregion
        }

        //Остановка потока
        private void StopSlaveThread()
        {
            #region Остановка потока

            //Если поток инициализирован и запущен
            if (threadSlave != null && threadSlave.IsAlive)
            {
                mbSlave.StopRTU();  //Останавливаем поток
                threadSlave.Join(); //Ждем его завершения
            }

            #endregion
        }

        //Инициализация потока
        private void InitThreads()
        {
            #region Инициализация потока

            //Создаем и заполняем Modbus Slave 3 данными из файлов
            mbSlave.CreateDataStore(3, ref AWAUS3, ref BWAUS3);
            //Создаем и заполняем Modbus Slave 4 данными из файлов
            mbSlave.CreateDataStore(4, ref AWAUS4, ref BWAUS4);
            //Передаем потоку функцию из класса ModbusService с номером порта
            threadSlave = new Thread(new ThreadStart(() => mbSlave.StartRTU("COM1")));
            //Помечаем поток как фоновый
            threadSlave.IsBackground = true;

            #endregion
        }

    }
}
