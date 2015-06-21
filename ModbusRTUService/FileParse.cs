using System;
using System.IO;
using System.Globalization;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Collections.Generic;


namespace ModbusRTUService
{
    /// <summary>
    /// Класс для обработки файлов с данными
    /// </summary>
    class FileParse : IFileParse
    {
        #region Считывание аналоговвых значений
        public ushort[] AWAUSParse(List<string> filesName)
        {
            
            //Устанавливаем разделитель для дроных значений
            CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.CurrencyDecimalSeparator = ".";

            //Номер считываемого файла
            int nFile = 0;
            
            //инициализация массива, который будем возвращать (количество файлов в списке * 2 байта для типа Float * количество элементов в файле)
            ushort[] AnalogOutput = new ushort[filesName.Count*2*100];

            
            //Цикл перебора файлов
            foreach (string fileName in filesName)
            {
                //инициализация счетчиков 
                // nLine - номер строки считывания
                // item - номер элемента в массиве
                // strTemp - строка считанная из файла
                int nLine = 0;
                string strTemp;
                int item = nFile * 200; //Номер файла

                
                try
                {
                    // открытие файла
                    using (StreamReader sr = new StreamReader(fileName))
                    {
                        // чтение строки
                        while ((strTemp = sr.ReadLine()) != null)
                        {
                            nLine++;
                            // если считали 100 элементов (4 первые строки заголовки чайла)
                            if (nLine > 104)
                            {
                                // выходим из цикла чтения (остальные строки в файле нам не нужны)
                                break;
                            }
                            // пропускаем первые 4 строки
                            else if (nLine > 4)
                            {
                                //считываем нужное значение как дробное и переводим в байтовый массив
                                byte[] buf = BitConverter.GetBytes(Single.Parse(strTemp.Substring(strTemp.IndexOf(",") + 1, 13).Trim().ToUpper(), NumberStyles.Any, ci));
                                //записываем побайтно в целочисленный массив в инверсированном порядке старшим регистром вперед
                                AnalogOutput[item] = BitConverter.ToUInt16(buf, 2);
                                AnalogOutput[item + 1] = BitConverter.ToUInt16(buf, 0);
                                //переводим счетчик массива на 2 элемента вперед
                                item += 2;
                            }
                            continue;
                        }

                    }

                }
                catch (Exception ex)
                {
                    //Если ошибка, то выводим в журнал событий
                    EventLog eventLog = new EventLog();
                    if (!EventLog.SourceExists("ModbusRTUService"))
                    {
                        EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService");
                    }
                    eventLog.Source = "ModbusRTUService";
                    eventLog.WriteEntry(fileName + " : " + ex.Message, EventLogEntryType.Error);
                }
                nFile++;
            }
            //возвращаем значения
            return AnalogOutput;
        }

        #endregion

        #region Считывание дискретных значений и преобразование их к аналоговым
        public ushort[] BWAUSParse(List<string> filesName)
        {
            
            //Номер считываемого файла
            int nFile = 0;
            
            //инициализация массива, который будем возвращать (количество преобразованных сигналов * количество файлов в списке)
            ushort[] DiscreteOutput = new ushort[6 * filesName.Count];

            foreach (string fileName in filesName)
            {
                //инициализация счетчиков 
                // nLine - номер строки считывания
                // bit - номер бита 
                // strTemp - строка считанная из файла
                // item - номер элемента в массиве
                int nLine = 0, bit = 0;
                string strTemp;
                int item = nFile * 6;

                
                //инициализация битового массива, который будем преобразовывать в целочисленное значение
                BitArray bits = new BitArray(16);

                try
                {
                    // открытие файла
                    using (StreamReader sr = new StreamReader(fileName))
                    {
                        // чтение строки
                        while ((strTemp = sr.ReadLine()) != null)
                        {
                            nLine++;
                            // если считали 100 элементов (4 первые строки заголовки чайла)
                            if (nLine > 104)
                            {
                                // выходим из цикла чтения
                                break;
                            }
                            // пропускаем первые 4 строки
                            else if (nLine > 4)
                            {
                                // заполняем массив из битов значениями из строк
                                bits[bit] = Convert.ToBoolean(Convert.ToByte(strTemp.Substring(strTemp.IndexOf(",") + 1, 1)));
                                bit++;
                                // если массив заполнен
                                if (bit == 16)
                                {
                                    //инициализируем массив для конвертации
                                    int[] buf = new int[1];
                                    // преобразуем в int
                                    bits.CopyTo(buf, 0);
                                    // переносим в массив для возвращения значения
                                    DiscreteOutput[item] = (ushort)buf[0];
                                    // обнуляем значения счетчика массива битов
                                    bit = 0;
                                    // изменяем счетчик возвращаемых значений
                                    item++;
                                }
                                // значений должно получится 6
                                if (item > 6)
                                    break;

                            }
                            continue;
                        }

                    }
                }
                catch (Exception ex)
                {
                    //Если ошибка, то выводим в журнал событий
                    EventLog eventLog = new EventLog();
                    if (!EventLog.SourceExists("ModbusRTUService"))
                    {
                        EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService");
                    }
                    eventLog.Source = "ModbusRTUService";
                    eventLog.WriteEntry(fileName + " : " + ex.Message, EventLogEntryType.Error);
                }
                nFile++;
            }

            //возвращаем значения
            return DiscreteOutput;
        }
        #endregion

        #region Считывание аналоговых значений другой системы
        public ushort[] ptoParse(string fileName)
        {
            //Устанавливаем разделитель для дроных значений
            CultureInfo ci = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.CurrencyDecimalSeparator = ".";

            //инициализация счетчиков 
            // chPlusCount - количество углов таблицы в файле ("+")
            // item - номер элемента в массиве
            // flag - флаг разрешения считывания данных из файла
            // line - строка считанная из файла
            int chPlusCount = 0, item = 0;
            bool flag = false;
            string line;

            //инициализация массива, который будем возвращать
            ushort[] AnalogOutput = new ushort[500];

            try
            {
                // открытие файла с нужной кодировкой
                using (StreamReader sr = new StreamReader(fileName, Encoding.GetEncoding("iso-8859-5")))
                {
                    // инициализация кодировок для перекодирования строки из файла
                    Encoding iso = Encoding.GetEncoding("ISO-8859-5");
                    Encoding win = Encoding.Default;

                    // чтение строки 
                    while ((line = sr.ReadLine()) != null)
                    {
                        // перекодировка байтов строки в нужную нам кодировку
                        byte[] isoBytes = iso.GetBytes(line);
                        byte[] winBytes = Encoding.Convert(iso, win, isoBytes);
                        char[] winChars = new char[win.GetCharCount(winBytes, 0, winBytes.Length)];
                        win.GetChars(winBytes, 0, winBytes.Length, winChars, 0);
                        line = new string(winChars);
                        /*
                         * Время отчета
                        if (line.StartsWith("Время за"))
                        {
                            line.Substring(line.IndexOf(".") - 2, 10);
                        }
                        */
                        // ищем начало таблицы
                        if (line.StartsWith("+"))
                        {
                            chPlusCount++;
                            if (chPlusCount == 2)
                            {
                                flag = true;
                                chPlusCount = 0;
                                continue;
                            }
                        }
                        // начало таблицы найдено и пропущены все заголовки в таблице
                        if (flag == true)
                        {
                            // если строка не пустая
                            if (line.Length != 0)
                            {
                                // разбивем строку на подстроки
                                string[] temp = line.Split(new char[] { '|' });
                                /*
                                 * temp[1] - порядковый номер в документе
                                 * temp[5] - наименование параметра
                                 * temp[7] - значение 
                                 * temp[8] - единицы измерения
                                dataGridView1.Rows.Add(temp[1].Trim(), temp[5].Trim(), temp[7].Trim(), temp[8].Trim());
                                */
                                //считываем нужное значение как дробное и переводим в байтовый массив
                                byte[] buf = BitConverter.GetBytes(Single.Parse(temp[7].Trim(), NumberStyles.Any, ci));
                                //записываем побайтно в целочисленный массив в инверсированном порядке старшим регистром вперед
                                AnalogOutput[item] = BitConverter.ToUInt16(buf, 2);
                                AnalogOutput[item + 1] = BitConverter.ToUInt16(buf, 0);
                                //переводим счетчик массива на 2 элемента вперед
                                item += 2;
                            }
                            else
                            // если строка пустая
                            {
                                flag = false;
                                chPlusCount = 0;
                            }
                            continue;
                        }

                    }

                }


            }
            catch (Exception ex)
            {
                //Если ошибка, то выводим в журнал событий
                EventLog eventLog = new EventLog();
                if (!EventLog.SourceExists("ModbusRTUService"))
                {
                    EventLog.CreateEventSource("ModbusRTUService", "ModbusRTUService");
                }
                eventLog.Source = "ModbusRTUService";
                eventLog.WriteEntry(fileName + " : " + ex.Message, EventLogEntryType.Error);
            }
            //возвращаем значения
            return AnalogOutput;
        }
        #endregion
    }
}
