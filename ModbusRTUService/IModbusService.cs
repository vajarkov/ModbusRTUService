using System;


namespace ModbusRTUService
{
    /// <summary>
    /// Интерфейс службы для инициализации, создания устройств, передачи данных устройствам и запуска устройства Modbus
    /// </summary>
    public interface IModbusService
    {
        /// <summary>
        /// Метод для передачи данных в хранилище Modbus-устройства
        /// </summary>
        /// <param name="slaveId">Номер утстройства Slave</param>
        /// <param name="AWAUS">Массив аналоговых значений</param>
        /// <param name="BWAUS">Массив дискретных значений, преобразованных в ushort[]</param>
        void CreateDataStore(byte slaveId, ref ushort[] AWAUS, ref ushort[] BWAUS);

        /// <summary>
        /// Метод для запуска устройства на нужном COM-порте
        /// </summary>
        /// <param name="comPort">выбранный COM-порт</param>
        void StartRTU(string comPort);

        /// <summary>
        /// Остановка службы устройтсва Modbus
        /// </summary>
        void StopRTU();
    }
}
