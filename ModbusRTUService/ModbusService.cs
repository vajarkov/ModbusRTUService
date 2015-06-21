using System;
using System.IO.Ports;
using System.Diagnostics;
using Modbus;
using Modbus.Data;
using Modbus.Device;
using System.Collections.Generic;

namespace ModbusRTUService
{

    /// <summary>
    /// Класс для работы с Modbus-устройством
    /// </summary>
    public class ModbusService : IModbusService
    {
        // Переменная для работы с устройством Modbus Slave
        private ModbusSlave slave;
        // Переменная для COM-порта
        private static SerialPort comPort;
        // Переменная для хранилища данных нескольких устройств Modbus Slave
        private static Dictionary<byte, DataStore> mapSlavesData = new Dictionary<byte, DataStore>();

        // Создание хранилища данных для Modbus Slave
        public void CreateDataStore(byte slaveId, ref ushort[] AWAUS, ref ushort[] BWAUS)
        {
            #region Создание хранилища

            //Создаем хранилище данных для ModbusSlave
            DataStore dataStore = DataStoreFactory.CreateDefaultDataStore();
            //Устанавливаем стартовый адрес для данных
            int nAddressMB = 1;

            #endregion

            #region Запись данных в хранилище

            //Записываем аналоговые значения
            foreach (ushort item in AWAUS)
            {
                dataStore.HoldingRegisters[nAddressMB] = item;
                nAddressMB++;
            }

            //Смещаем адрес для дискретных значений
            nAddressMB = 1001;

            //Записываем дискретные значения
            foreach (ushort item in BWAUS)
            {
                dataStore.HoldingRegisters[nAddressMB] = item;
                nAddressMB++;
            }

            #endregion

            #region Добавление хранилища в общий массив

            mapSlavesData.Add(slaveId, dataStore);

            #endregion

        }

        //Запуск Modbus-устройства
        public void StartRTU(string portName)
        {
            #region Создание и запуск устройства
            try
            {
                #region Инициализируем и открываем COM-порт
                //Создаем COM-порт
                using (comPort = new SerialPort(portName, 19200, Parity.None, 8, StopBits.One ))
                {
                    
                    //Открываем порт, если закрыт
                    if (!comPort.IsOpen)
                        comPort.Open();
                #endregion

                    #region Создание Modbus-устройства и его запуск
                    //Создаем устройство
                    slave = ModbusSerialSlave.CreateRtu(mapSlavesData, comPort);

                    //Запускаем устройства
                    slave.Listen();
                    #endregion
                }
            }
            catch (Exception ex)
            {
                //Если ошибка пишем в журнал
                EventLog eventLog = new EventLog();
                if (!EventLog.SourceExists("ModbusRTUService"))
                {
                    EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService");
                }
                eventLog.Source = "ModbusRTUService";
                eventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            #endregion
        }

        //Остановка Modbus-устройства
        public void StopRTU()
        {
            #region Отсновка устройства

            //Если устройство создано
            if (slave != null)
            {
                //Очистить общее хранилище данных
                mapSlavesData.Clear();
                //Послать флаг остановки цикла чтения
                slave.stop = true;

            }

            #endregion
        }
    }
}
