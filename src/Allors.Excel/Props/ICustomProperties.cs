namespace Allors.Excel
{
    using System;

    public interface ICustomProperties
    {
        bool? GetBoolean(string name);

        DateTime? GetDate(string name);

        double? GetFloat(string name);

        long? GetNumber(string name);

        string? GetString(string name);

        void SetBoolean(string name, bool? value);

        void SetDate(string name, DateTime? value);

        void SetFloat(string name, double? value);

        void SetNumber(string name, long? value);

        void SetString(string name, string? value);

        void Remove(string name);
    }
}
