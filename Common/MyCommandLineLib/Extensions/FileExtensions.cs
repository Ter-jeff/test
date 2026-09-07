using System;

using MyCommandLineLib.Enums;

namespace MyCommandLineLib.Extensions
{
    public static class FileExtensions
    {
        public static string GetExtension(this EnumFileExtension enumFileExtension)
        {
            string extStr = enumFileExtension switch
            {
                EnumFileExtension.Ini => ".ini",
                EnumFileExtension.Txt => ".txt",
                EnumFileExtension.Log => ".log",
                EnumFileExtension.Xml => ".xml",
                _ => throw new ArgumentOutOfRangeException(nameof(enumFileExtension), enumFileExtension, null),
            };
            return extStr;
        }
    }
}
