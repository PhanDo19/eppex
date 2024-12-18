﻿using OfficeOpenXml;
using System.Reflection;

namespace changeExcel.Utils
{
    static class EPPlusExtension
    {
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet excelWorksheet) where T : new()
        {
            Func<CustomAttributeData, bool> columnOnly = (x => x.AttributeType == typeof(Column));

            var columns = typeof(T).GetProperties()
                .Where(x => x.CustomAttributes.Any(columnOnly))
                .Select(x => new
                {
                    Property = x,
                    Column = x.GetCustomAttributes<Column>().First().ColumnIndex,
                    Required = x.GetCustomAttributes<Column>().First().IsRequired
                }).ToList();

            var rows = excelWorksheet.Cells
                .Select(x => x.Start.Row)
                .Distinct()
                .OrderBy(x => x);

            var collection = new List<T>();
             var code = string.Empty;
            foreach (var row in rows.Skip(3))
            {
                var newObject = new T();
                foreach (var col in columns)
                {
                    try
                    {
                        ExcelRange currentValue = excelWorksheet.Cells[row, col.Column];
                       
                        if (currentValue.Value == null)
                        {
                            // if condition is true, continue to next iteration
                            if (col.Required)
                            {
                                break;
                            }
                            col.Property.SetValue(newObject, null);
                        }
                        else if (col.Property.PropertyType == typeof(Int32))
                        {
                            var value = currentValue.GetValue<int>();
                            col.Property.SetValue(newObject, value);
                        }
                        else if (col.Property.PropertyType == typeof(decimal))
                        {
                            var stringValue = currentValue.GetValue<string>();
                            if (stringValue.Contains("%"))
                            {
                                var numericValue = decimal.Parse(stringValue.TrimEnd('%'));
                                col.Property.SetValue(newObject, numericValue);
                            }
                            else
                            {
                                var value = currentValue.GetValue<decimal>();
                                col.Property.SetValue(newObject, currentValue.GetValue<decimal>());
                            }
                        }
                        else if (col.Property.PropertyType == typeof(DateTime))
                        {
                            var value = currentValue.GetValue<DateTime>();
                            col.Property.SetValue(newObject, currentValue.GetValue<DateTime>());
                        }
                        else if (col.Property.PropertyType == typeof(bool))
                        {
                            var value = currentValue.GetValue<bool>();
                            col.Property.SetValue(newObject, currentValue.GetValue<bool>());
                        }
                        else if (col.Property.PropertyType == typeof(string))
                        {
                            var value = currentValue.GetValue<string>();
                            col.Property.SetValue(newObject, value);
                            if (col.Required)
                            {
                               code = value;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Equals("required_data"))
                        {
                            throw new Exception("Lỗi đơn sản phẩm : " + code + " column: " + col.Property.Name);
                        }

                        throw new Exception("Lỗi đơn sản phẩm : " + code + " column: " + col.Property.Name);
                    }
                }
                // value has property name is Name null or empty continue to next iteration
                if (string.IsNullOrEmpty(newObject.GetType().GetProperty("Name")?.GetValue(newObject)?.ToString())) continue;
                collection.Add(newObject);
            }
            return collection;
        }
    }
}
