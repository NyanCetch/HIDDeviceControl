using System;
using System.Linq;

namespace HIDDeviceControl
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var devices = HIDDevices.Enumerate().ToArray();
            Console.WriteLine("Число подключенных устройств: " + devices.Length);

            var success = HIDDevices.EnableDevice(new HIDDevice{InstanceID = @"USB\VID_09DA&PID_1686&MI_01\6&34cedc73&0&0001"}, out var resultMsg);
            Console.WriteLine($"Операция {(success ? "" : "не")} выполнена");
            if (!success)
                Console.WriteLine(resultMsg);
        }
    }
}    