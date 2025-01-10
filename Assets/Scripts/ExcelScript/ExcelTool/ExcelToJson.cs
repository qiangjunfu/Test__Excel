#if UNITY_EDITOR
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using UnityEditor;


public class ExcelToJson 
{


    [MenuItem("FQJ/ExcelTool/Excel_To_Class")]
    public static void XLSX_To_Class() 
    {
        List<string> fileNameList = new List<string>();
        string path = string.Format("{0}/{1}/{2}", UnityEngine.Application.dataPath, "StreamingAssets", "Excels");
        GetFiles(path, "xlsx", ref fileNameList);

        for (int i = 0; i < fileNameList.Count; i++)
        {
            string xlsxPath = fileNameList[i];

            // xlsx工作表名名
            string SheetNames = Path.GetFileNameWithoutExtension(xlsxPath);
            UnityEngine.Debug.Log("工作表名:  " + SheetNames);

            FileStream stream = null;
            try
            {
                stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException e)
            {
                UnityEngine.Debug.LogErrorFormat("读取文件错误:   " + e);
            }
            if (stream == null)
            {
                UnityEngine.Debug.LogErrorFormat("stream == null  ");
                return;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            // 读取第一张工资表
            DataTable dataTable = result.Tables[0];
            CreateXLSXClass(SheetNames, dataTable);
            AssetDatabase.Refresh();


            excelReader.Close();
        }
    }


    [MenuItem("FQJ/ExcelTool/Excel_To_Json")]
    public static void XLSX_To_Json()  
    {
        List<string> fileNameList = new List<string>();
        string path = string.Format("{0}/{1}/{2}", UnityEngine.Application.dataPath, "StreamingAssets", "Excels");
        GetFiles(path, "xlsx", ref fileNameList);

        for (int i = 0; i < fileNameList.Count; i++)
        {
            string xlsxPath = fileNameList[i];

            // xlsx工作表名名
            string SheetNames = Path.GetFileNameWithoutExtension(xlsxPath);
            UnityEngine.Debug.Log("工作表名:  " + SheetNames);

            FileStream stream = null;
            try
            {
                stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException e)
            {
                UnityEngine.Debug.LogErrorFormat("读取文件错误:   " + e);
            }
            if (stream == null)
            {
                UnityEngine.Debug.LogErrorFormat("stream == null  ");
                return;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            // 读取第一张工资表
            Type t = Type.GetType(SheetNames);
            DataTable dataTable = result.Tables[0];

            string directoryPath = string.Format("{0}/{1}", UnityEngine.Application.streamingAssetsPath, "Jsons");
            string savePath = string.Format("{0}/{1}/{2}.json", UnityEngine.Application.streamingAssetsPath, "Jsons", SheetNames);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath); // 如果目录不存在，则创建
                UnityEngine.Debug.Log("目录不存在，已创建：" + directoryPath);
            }
            AssetDatabase.Refresh();

            ReadSingleSheet(t, dataTable, savePath );
            UnityEngine.Debug.Log("保存完成 !！!  " + savePath);
            AssetDatabase.Refresh();


            excelReader.Close();
        }
    }



    /// <summary>
    /// 读取一个工作表的数据
    /// </summary>
    /// <param name="type">要转换的struct或class类型</param>
    /// <param name="dataTable">读取的工作表数据</param>
    /// <param name="jsonPath">存储路径</param>
    private static void ReadSingleSheet(Type type, DataTable dataTable, string jsonPath)
    {
        int rows = dataTable.Rows.Count;
        int Columns = dataTable.Columns.Count;
        // 工作表的行数据
        DataRowCollection collect = dataTable.Rows;
        // xlsx对应的数据字段，规定是第二行
        string[] jsonFileds = new string[Columns];
        // 要保存成Json的obj
        List<object> objsToSave = new List<object>();
        for (int i = 0; i < Columns; i++)
        {
            jsonFileds[i] = collect[1][i].ToString();
        }

        // 从第三行开始
        for (int i = 3; i < rows; i++)
        {
            // 生成一个实例
            Object objIns = type.Assembly.CreateInstance(type.ToString());

            for (int j = 0; j < Columns; j++)
            {
                // 获取字段
                FieldInfo field = type.GetField(jsonFileds[j]);
                if (field != null)
                {
                    object value = null;
                    try // 赋值
                    {
                        UnityEngine.Debug.Log("赋值:  " + collect[i] + "  " + collect[i][j] + "   type: " + field.FieldType);
                        value = Convert.ChangeType(collect[i][j], field.FieldType);
                    }
                    catch (InvalidCastException e)
                    {
                        UnityEngine.Debug.LogWarningFormat("赋值错误:  " + e.Message + "\n" + collect[i] + "  " + collect[i][j] + "   type: " + field.FieldType);

                        // ------ 处理不同类型的数组 ------
                        if (field.FieldType.ToString() == "System.Int32[]")
                        {
                            //UnityEngine.Debug.LogError("System.Int32[]");
                            string str = collect[i][j].ToString();
                            string[] strs = str.Split(',');
                            int[] ints = new int[strs.Length];
                            for (int k = 0; k < strs.Length; k++)
                            {
                                ints[k] = int.Parse(strs[k]);
                            }
                            value = ints;
                        }
                        else if (field.FieldType.ToString() == "System.Single[]")
                        {
                            //UnityEngine.Debug.LogError("System.Single[]");
                            string str = collect[i][j].ToString();
                            string[] strs = str.Split(',');
                            System.Single[] singles = new System.Single[strs.Length];
                            for (int k = 0; k < strs.Length; k++)
                            {
                                singles[k] = System.Single.Parse(strs[k]);
                            }
                            value = singles;
                        }
                        else if (field.FieldType.ToString() == "System.String[]")
                        {
                            //UnityEngine.Debug.LogError("System.String[]");
                            string str = collect[i][j].ToString();
                            string[] strs = str.Split(',');
                            String[] strings = new String[strs.Length];
                            for (int k = 0; k < strs.Length; k++)
                            {
                                strings[k] = strs[k];
                            }
                            value = strings;
                        }
                        else if (field.FieldType.ToString() == "System.Boolean[]")
                        {
                            // ------ 处理 bool[] ------
                            string str = collect[i][j].ToString();
                            string[] strs = str.Split(',');
                            bool[] bools = new bool[strs.Length];
                            for (int k = 0; k < strs.Length; k++)
                            {
                                bools[k] = bool.Parse(strs[k].Trim()); // 去除多余空格，并解析
                            }
                            value = bools;
                        }
                    }

                    field.SetValue(objIns, value);
                }
                else
                {
                    UnityEngine.Debug.LogFormat("有无法识别的字符串：{0}", jsonFileds[j]);
                }
            }
            objsToSave.Add(objIns);
        }
        // 保存为Json
        string content = Newtonsoft.Json.JsonConvert.SerializeObject(objsToSave);
        SaveFile(content, jsonPath);
    }


    private static void GetFiles(string path, string suffix, ref List<string> fileNameList, bool isSubcatalog = false)
    {
        string filename;
        DirectoryInfo dir = new DirectoryInfo(path);
        FileInfo[] file = dir.GetFiles();
        //如需遍历子文件夹时需要使用  
        DirectoryInfo[] dii = dir.GetDirectories();

        foreach (FileInfo f in file)
        {
            filename = f.FullName;//拿到了文件的完整路径
            if (filename.EndsWith(suffix))//判断文件后缀，并获取指定格式的文件全路径增添至fileList  
            {
                fileNameList.Add(filename);
                //UnityEngine. Debug.Log(filename);
            }
        }
        //获取子文件夹内的文件列表，递归遍历
        if (isSubcatalog)
        {
            foreach (DirectoryInfo d in dii)
            {
                GetFiles(d.FullName, "", ref fileNameList, false);
            }
        }

        return;
    }


    private static void SaveFile(string content, string jsonPath)
    {
        StreamWriter streamWriter;
        FileStream fileStream;
        if (File.Exists(jsonPath))
        {
            File.Delete(jsonPath);
        }
        fileStream = new FileStream(jsonPath, FileMode.Create);
        streamWriter = new StreamWriter(fileStream);
        streamWriter.Write(content);
        streamWriter.Close();
    }



    #region 
    private static void CreateXLSXClass(string className, DataTable dataTable)
    {
        int rows = dataTable.Rows.Count;
        int Columns = dataTable.Columns.Count;
        // 工作表的行数据
        DataRowCollection collect = dataTable.Rows;

        List<string> ziduan = new List<string>();
        List<string> name = new List<string>();
        List<string> annotation = new List<string>();
        for (int i = 0; i < Columns; i++)
        {
            UnityEngine.Debug.LogFormat("{3}类  {0}  注释: {4}  字段: {2}   名称: {1} ", i, collect[1][i].ToString(), collect[2][i].ToString(), className , collect[0][i].ToString());
            ziduan.Add(collect[2][i].ToString());
            name.Add(collect[1][i].ToString());
            annotation.Add(collect[0][i].ToString());
        }
        GenerateScript(className, ziduan, name , annotation);
    }


    static void GenerateScript(string scriptName, List<string> ziduan, List<string> name, List<string> annotation)
    {
        if (scriptName != null)
        {
            string templatePath = "Assets/Scripts/ExcelScript/ExcelTool/DataScriptTemplate.cs";
            string spawnPath = "Assets/Scripts/ExcelScript/";
            string scriptPath = spawnPath + scriptName + ".cs";

            string templateContent = File.ReadAllText(templatePath);
            //替换类名
            templateContent = templateContent.Replace("DataScriptTemplate", scriptName);
            templateContent = templateContent.Replace("}", " ");
            // 往模板类添加属性字段
            string str = "";
            for (int i = 0; i < ziduan.Count; i++)
            {
                str += "    /// <summary>\n" + "    ///  " + annotation[i] + "\n    /// </summary>\n"; 
                str += "    public " + ziduan[i] + "  " + name[i] + " ; \n";
            }
            str += " \n  }";
            templateContent += str;



            string directoryPath = Path.GetDirectoryName(scriptPath); // 获取文件目录路径
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath); // 如果目录不存在，则创建
                UnityEngine.Debug.Log("目录不存在，已创建：" + directoryPath);
            }
            //AssetDatabase.Refresh();

            try
            {
                // 将内容写入文件
                File.WriteAllText(scriptPath, templateContent, System.Text.Encoding.UTF8);
                UnityEngine.Debug.Log("文件成功写入：" + scriptPath);
            }
            catch (Exception ex)
            {
                UnityEngine.Debug.LogError("写入文件失败：" + ex.Message);
            }

            ////将组织好的内容写入文件
            //File.WriteAllText(scriptPath, templateContent, System.Text.Encoding.UTF8);
            ////刷新一下资源，不然创建好文件后第一时间不会显示
            //AssetDatabase.Refresh();
        }
        else
        {
            UnityEngine.Debug.LogWarning("文件名不能为空");
        }
    }

    #endregion 
}
#endif