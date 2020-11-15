using System;

namespace CrmCodeGenerator.VSPackage
{
    public class MapperEventArgs : EventArgs
    {
        public string Message { get; set; }
        public string MessageExtended { get; set; }
    }
}
