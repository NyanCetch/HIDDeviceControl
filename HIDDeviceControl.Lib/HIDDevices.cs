using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Newtonsoft.Json;

namespace HIDDeviceControl
{
    /// <summary>
    /// Вспомогательный класс для получения информации о HID-устройствах и включения/выключения их
    /// Работает на базе Powershell и PnPUtil
    /// </summary>
    public static class HIDDevices
    {
        private const string PNPUTIL_PATH = @"C:\Windows\SysNative\pnputil.exe";
        
        /// <summary>
        /// Выдает список подключенных HID-устройств
        /// </summary>
        /// <returns>Найденные устройства</returns>
        public static IEnumerable<HIDDevice> Enumerate() => Enumerate(out _, out _);
        
        /// <summary>
        /// Выдает список подключенных HID-устройств
        /// </summary>
        /// <param name="messages">Сообщения в Powershell</param>
        /// <param name="errors">Ошибки в Powershell</param>
        /// <returns>Найденные устройства</returns>
        public static IEnumerable<HIDDevice> Enumerate(out string[] messages, out string[] errors)
        {
            errors = Array.Empty<string>();
            messages = Array.Empty<string>();
            
            var script = @"((& " + PNPUTIL_PATH +
                         @" /enum-devices /class HIDClass /connected | Select-Object -Skip 2) | Select-String -Pattern 'GUID.+:' -Context 3,4) | 
                ForEach {
                    $removeRe = '.*:\s*'
                    [PSCustomObject]@{
                        InstanceID = $PSItem.Context.PreContext[0] -replace $removeRe
                        Description = $PSItem.Context.PreContext[1] -replace $removeRe
                        ClassName = $PSItem.Context.PreContext[2] -replace $removeRe
                        ClassGUID = ($PSItem | Select-String -Pattern 'GUID.+:') -replace $removeRe
                        ProviderName = $PSItem.Context.PostContext[0] -replace $removeRe
                        Status = $PSItem.Context.PostContext[1] -replace $removeRe
                        DriverName = $PSItem.Context.PostContext[2] -replace $removeRe
                    }
                }
            ";

            // Собираем результаты, если они есть, в виде JSON
            // Note: это нужно чтобы починить кодировку
            var deviceJsonList = new List<string>();
            using (var ps = PowerShell.Create())
            {
                ps.AddScript(script);
                var results = ps.Invoke();
                
                if (results.Count > 0)
                {
                    foreach (var res in results)
                    {
                        var propDict = res.Properties.ToDictionary(p => p.Name, p => p.Value);
                        var jsonStr = JsonConvert.SerializeObject(propDict, Formatting.Indented);
                        deviceJsonList.Add(RepairEncoding("CP866", "windows-1251", jsonStr));
                    }
                }
                
                if (ps.Streams.Error.Count > 0)
                {
                    errors = ps.Streams.Error.Select(err => err.ToString()).ToArray();
                }
                
                if (ps.Streams.Debug.Count > 0)
                {
                    messages = ps.Streams.Debug.Select(msg => msg.ToString()).ToArray();
                }
            }
            
            var deviceList = new List<HIDDevice>();
            foreach (var deviceJson in deviceJsonList)
            {
                var device = JsonConvert.DeserializeObject<HIDDevice>(deviceJson);
                deviceList.Add(device);
            }

            return deviceList;
        }

        /// <summary>
        /// Включает HID-устройство
        /// </summary>
        /// <param name="device">Устройство</param>
        /// <param name="resultMsg">Статус выполнения</param>
        /// <returns>true - если операция выполнена, иначе - false</returns>
        public static bool EnableDevice(HIDDevice device, out string resultMsg)
        {
            return SetDeviceActive(device, true, out resultMsg);
        }

        /// <summary>
        /// Выключает HID-устройство
        /// </summary>
        /// <param name="device">Устройство</param>
        /// <param name="resultMsg">Статус выполнения</param>
        /// <returns>true - если операция выполнена, иначе - false</returns>
        public static bool DisableDevice(HIDDevice device, out string resultMsg)
        {
            return SetDeviceActive(device, false, out resultMsg);
        }

        /// <summary>
        /// Включает или выключает HID-устройство
        /// </summary>
        /// <param name="device">Устройство</param>
        /// <param name="active">Флаг включить или нет</param>
        /// <param name="resultMsg">Статус выполнения</param>
        /// <returns>true - если операция выполнена, иначе - false</returns>
        private static bool SetDeviceActive(HIDDevice device, bool active, out string resultMsg)
        {
            /*
             Для управления устройствами используется команда pnputil /<action>-device "<InstanceID>".
             При выполнении выводит 5 строк: первые 2 - шапка, 5 - пустая строка, 4 - статус выполнения 
             */
            
            var isSuccess = false;
            resultMsg = string.Empty;
            
            var commandFormat = "({0} /{1}-device \"{2}\") | Select-Object -Index 3";
            var action = active ? "enable" : "disable";
            var command = string.Format(commandFormat, PNPUTIL_PATH, action, device.InstanceID);

            using (var ps = PowerShell.Create())
            {
                ps.AddScript(command);
                var results = ps.Invoke();

                if (results.Count > 0)
                {
                    var statusLine = RepairEncoding("CP866", "windows-1251", results[0].ToString());
                    isSuccess = statusLine.Contains("успешно");

                    resultMsg = statusLine;
                }
            }

            return isSuccess;
        }
        
        /// <summary>
        /// Исправляет кодировку строки
        /// </summary>
        /// <param name="from">В чем была закодирована строка</param>
        /// <param name="to">В чем она должна отображатся</param>
        /// <param name="str">Строка, закодированная в кодировке from</param>
        private static string RepairEncoding(string from, string to, string str)
        {
            var fromEnc = Encoding.GetEncoding(from);
            var toEnc = Encoding.GetEncoding(to);

            return toEnc.GetString(fromEnc.GetBytes(str));
        }
    }
}