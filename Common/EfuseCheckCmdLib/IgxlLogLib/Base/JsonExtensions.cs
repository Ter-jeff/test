using System;
using System.Linq;

using Newtonsoft.Json.Serialization;

namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    internal static class JsonExtensions
    {
        public static string[] PropertyNames(this IContractResolver contractResolver, Type type)
        {
            ArgumentNullException.ThrowIfNull(contractResolver);
            ArgumentNullException.ThrowIfNull(type);

            if (contractResolver.ResolveContract(type) is not JsonObjectContract contract)
            {
                return [];
            }

            return [.. contract.Properties.Where(p => !p.Ignored).Select(p => p.PropertyName!)];
        }
    }
}
