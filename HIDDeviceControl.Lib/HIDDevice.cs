using System;

namespace HIDDeviceControl
{
    public class HIDDevice
    {
        public string InstanceID { get; set; }
        public string Description { get; set; }
        public string ClassName {get; set;}
        public string ClassGUID {get; set;}
        public string ProviderName {get; set;}
        public string Status {get; set;}
        public string DriverName {get; set;}

        public HIDDevice()
        {
            
        }
    }
}