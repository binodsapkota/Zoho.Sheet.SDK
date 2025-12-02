using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Utilities
{
    public static class SheetRecordMapper
    {
        // Map a single object to SheetRecord
        public static SheetRecord FromObject<T>(T obj, int rowIndex = 0)
        {
            if (obj == null) throw new ArgumentNullException(nameof(obj));

            var record = new SheetRecord { RowIndex = rowIndex };
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in props)
            {
                var value = prop.GetValue(obj);
                record.Data[prop.Name] = value;
            }

            return record;
        }

        // Map a list of objects to SheetRecords
        public static List<SheetRecord> FromObjects<T>(List<T> objects)
        {
            var list = new List<SheetRecord>();
            if (objects == null || objects.Count == 0) return list;

            int index = 0;
            foreach (var obj in objects)
            {
                list.Add(FromObject(obj, ++index));
            }

            return list;
        }
        // NEW: SheetRecord -> Object (TaskItem, or any T)
        public static T ToObject<T>(SheetRecord record) where T : new()
        {
            if (record == null) throw new ArgumentNullException(nameof(record));

            var obj = new T();
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in props)
            {
                if (record.Data.ContainsKey(prop.Name) && record.Data[prop.Name] != null)
                {
                    var value = record.Data[prop.Name];

                    // Convert value to property type safely
                    try
                    {
                        var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                        var convertedValue = Convert.ChangeType(value, targetType);
                        prop.SetValue(obj, convertedValue);
                    }
                    catch
                    {
                        // Ignore conversion errors, or optionally log them
                    }
                }
            }

            return obj;
        }

        // Convert list of SheetRecords to list of T
        public static List<T> ToObjects<T>(List<SheetRecord> records) where T : new()
        {
            var list = new List<T>();
            if (records == null || records.Count == 0) return list;

            foreach (var record in records)
            {
                list.Add(ToObject<T>(record));
            }

            return list;
        }
    }
    public class SheetRecord
    {
        public int RowIndex { get; set; } = 0; // Optional: can be used for updates
        public Dictionary<string, object> Data { get; set; } = new Dictionary<string, object>();
    }
}
