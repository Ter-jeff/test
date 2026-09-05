using System;
using System.IO;
using System.Security.Cryptography;

namespace Automation.GenerateIgxl.PostAction.VersionTracing
{
    public static class HashCalculator
    {
        public static string CalculateMd5(string filePath)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(filePath);
                using (var md5 = MD5.Create())
                {
                    byte[] hash = md5.ComputeHash(fileBytes);
                    return BitConverter.ToString(hash).Replace("-", "").ToLower();
                }
            }
            catch (Exception ex)
            {
                return $"Exception {ex.GetType().Name} for {filePath}: {ex.Message}";
            }
        }
    }
}
