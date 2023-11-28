using System.CommandLine;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

class Program {
    static async Task<int> Main(string[] args) {
        var ExcelOption = new Option<FileInfo?>(
            name: "--excel",
            description: "excel文件路径");
        var OutPathOption = new Option<DirectoryInfo?>(
            name: "--out_path",
            description: "json输出目录");
        var CsPathOption = new Option<DirectoryInfo?>(
            name: "--cs_path",
            description: "cs输出目录");
        var NamespaceOption = new Option<string>(
            name: "--namespace",
            description: "命名空间");

        var ExportCommand = new Command("export", "导出配置表") { ExcelOption, OutPathOption, CsPathOption, NamespaceOption };
        ExportCommand.SetHandler(ExportExcel, ExcelOption, OutPathOption, CsPathOption, NamespaceOption);
        
        var RootCommand = new RootCommand("导表工具");
        RootCommand.AddCommand(ExportCommand);
        return await RootCommand.InvokeAsync(args);
    }

    static void ExportExcel(FileInfo ExcelFile, DirectoryInfo OutDir, DirectoryInfo CsDir, string Namespace) {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var Pck = new ExcelPackage(ExcelFile);
        var Ws = Pck.Workbook.Worksheets[0];
    
        var Class = CollectClass(Ws);
        var Data = CollectData(Ws);

        var CsCode = "";
        if (string.IsNullOrEmpty(Namespace)) {
            CsCode = CreateClass(Class);
        }
        else {
            CsCode = @$"namespace {Namespace} {{
{CreateClass(Class)}
}}";
        }
        File.WriteAllText(Path.Combine(CsDir.FullName, $"{Class["__ClassName"]}.cs"), CsCode);
        File.WriteAllText(Path.Combine(OutDir.FullName, $"{Class["__ClassName"]}.json"), Data.ToString(Formatting.None));
        Console.WriteLine("导出成功");
    }
    
    static string CreateClass(JObject Class) {
        var FieldsCode = "";
        foreach (var Item in Class) {
            if (Item.Key == "__ClassName") {
                continue;
            }

            FieldsCode += @$"
    /// <summary>
    /// {Item.Value!["Manual"]}
    /// </summary>
    public {Item.Value!["Type"]} {Item.Key};
";
        }
        var ClassCode = $@"public class {Class["__ClassName"]} {{{FieldsCode}
}}";
        return ClassCode;
    }

    static JObject CollectData(ExcelWorksheet Worksheet) {
        var minColumnNum = Worksheet.Dimension.Start.Column;//工作区开始列
        var maxColumnNum = Worksheet.Dimension.End.Column; //工作区结束列
        var minRowNum = Worksheet.Dimension.Start.Row; //工作区开始行号
        var maxRowNum = Worksheet.Dimension.End.Row; //工作区结束行号
        var FieldInfoByCol = new JObject();
        for (int j = minColumnNum; j <= maxColumnNum; j++) {
            var FieldInfo = new JObject {
                { "Manual", Worksheet.Cells[minRowNum + 0, j].GetValue<string>() },
                { "FieldName", Worksheet.Cells[minRowNum + 1, j].GetValue<string>() },
                { "Type", Worksheet.Cells[minRowNum + 2, j].GetValue<string>() },
            };
            FieldInfoByCol.Add(j.ToString(), FieldInfo);
        }

        var Data = new JObject();
        for (int i = minRowNum + 3; i <= maxRowNum; i++) {
            var Item = new JObject();
            for (int j = minColumnNum; j <= maxColumnNum; j++) {
                var FieldInfo = FieldInfoByCol[j.ToString()]!;
                var FieldName = FieldInfo["FieldName"]!.Value<string>()!;
                var FieldType = FieldInfo["Type"]!.Value<string>()!;
                var FieldValue = Worksheet.Cells[i, j].GetValue<string>()!.Replace(" ", "");

                if (!FieldType.Contains("[]")) { // 不是数组
                    try {
                        Item.Add(FieldName, JsonConvert.DeserializeObject<JToken>(FieldValue));
                    }
                    catch (Exception e) {
                        // Console.WriteLine(e);
                        Item.Add(FieldName, FieldValue);
                    }
                }
                else { // 是数组
                    try {
                        var Jsons = FieldValue.Split("\n");
                        var ArrayData = new JArray();
                        foreach (var Json in Jsons) {
                            ArrayData.Add(JsonConvert.DeserializeObject<JToken>(Json)!);
                        }
                        Item.Add(FieldName, ArrayData);
                    }
                    catch (Exception e) {
                        // Console.WriteLine(e);
                        Item.Add(FieldName, new JArray(FieldValue.Split(",")));
                    }
                }
                
            }
            Data.Add(Item["Id"]!.Value<string>()!, Item);
        }

        return Data;
    }

    static JObject CollectClass(ExcelWorksheet Worksheet) {
        var minColumnNum = Worksheet.Dimension.Start.Column;//工作区开始列
        var maxColumnNum = Worksheet.Dimension.End.Column; //工作区结束列
        var minRowNum = Worksheet.Dimension.Start.Row; //工作区开始行号
        var Class = new JObject { { "__ClassName", Worksheet.Name } };
        for (int j = minColumnNum; j <= maxColumnNum; j++) {
            var FieldManual = Worksheet.Cells[minRowNum + 0, j].GetValue<string>()!;
            var FieldName   = Worksheet.Cells[minRowNum + 1, j].GetValue<string>()!;
            var FieldType   = Worksheet.Cells[minRowNum + 2, j].GetValue<string>()!;
            Class.Add(FieldName, new JObject {
                { "Manual", FieldManual },
                { "Type", FieldType },
            });
        }
        return Class;
    }
}