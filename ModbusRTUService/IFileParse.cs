using System;
using System.Collections.Generic;


namespace ModbusRTUService
{
    /// <summary>
    /// Интерфейс для обработки файлов
    /// </summary>
    interface IFileParse
    {
        /// <summary>
        /// Метод обработки файла с аналоговыми значениями
        /// </summary>
        /// <param name="filesName">Путь и имена файлов</param>
        /// <returns>Массив данных ushort[]</returns>
        ushort[] AWAUSParse(List<string> filesName);

        /// <summary>
        /// Метод обработки файла с дискретными значениями
        /// </summary>
        /// <param name="fileName">Путь и имя файла</param>
        /// <returns>Массив данных ushort[]</returns>
        ushort[] BWAUSParse(List<string> filesName);

        
        /// <summary>
        /// Метод обработки файла с аналоговыми значениями с ПТО
        /// </summary>
        /// <param name="fileName">Путь и имя файла</param>
        /// <returns>Массив данных ushort[]</returns>
        ushort[] ptoParse(string fileName);
    }
}
