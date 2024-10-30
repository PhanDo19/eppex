using DocumentFormat.OpenXml.VariantTypes;
using OfficeOpenXml;
using System.Reflection;

namespace changeExcel.Utils
{
    public class EPPlusExtension
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
            //foreach(var row in rows.Skip(3))
            //{
            //       foreach (var col in columns)
            //    {
            //        var currentValue = excelWorksheet.Cells[row, col.Column];
            //        if (currentValue.Value == null)
            //        {
            //            if (col.Required)
            //            {
            //                throw new Exception("required_data");
            //            }

            //            col.Property.SetValue(new T(), null);
            //        }
            //        else if (col.Property.PropertyType == typeof(Int32))
            //        {
            //            col.Property.SetValue(new T(), currentValue.GetValue<int>());
            //        }
            //        else if (col.Property.PropertyType == typeof(double))
            //        {
            //            col.Property.SetValue(new T(), currentValue.GetValue<double>());
            //        }
            //        else if (col.Property.PropertyType == typeof(DateTime))
            //        {
            //            col.Property.SetValue(new T(), currentValue.GetValue<DateTime>());
            //        }
            //        else if (col.Property.PropertyType == typeof(bool))
            //        {
            //            col.Property.SetValue(new T(), currentValue.GetValue<bool>());
            //        }
            //        else if (col.Property.PropertyType == typeof(string))
            //        {
            //            col.Property.SetValue(new T(), currentValue.GetValue<string>());
            //        }
            //    }
            //}

            IEnumerable<T> collection = rows.Skip(3)
                .Select(row =>
                {
                    var newObject = new T();

                    columns.ForEach(col =>
                    {
                        try
                        {
                            ExcelRange currentValue = excelWorksheet.Cells[row, col.Column];

                            if (currentValue.Value == null)
                            {
                                if (col.Required)
                                {
                                    throw new Exception("required_data");
                                }

                                col.Property.SetValue(newObject, null);
                            }
                            else if (col.Property.PropertyType == typeof(Int32))
                            {
                                var value = currentValue.GetValue<int>();
                                col.Property.SetValue(newObject, value);
                            }
                            else if (col.Property.PropertyType == typeof(double))
                            {
                                var value = currentValue.GetValue<double>();
                                col.Property.SetValue(newObject, currentValue.GetValue<double>());
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
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Equals("required_data"))
                            {
                                throw new Exception("Null data error at row " + row + " column: " + col.Property.Name);
                            }

                            throw new Exception("Conversion error at row: " + row + " column: " + col.Property.Name);
                        }
                    });

                    return newObject;
                });

            return collection;
        }
    }
}
