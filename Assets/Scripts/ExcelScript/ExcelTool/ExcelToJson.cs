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

            // xlsx����������
            string SheetNames = Path.GetFileNameWithoutExtension(xlsxPath);
            UnityEngine.Debug.Log("��������:  " + SheetNames);

            FileStream stream = null;
            try
            {
                stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException e)
            {
                UnityEngine.Debug.LogErrorFormat("��ȡ�ļ�����:   " + e);
            }
            if (stream == null)
            {
                UnityEngine.Debug.LogErrorFormat("stream == null  ");
                return;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            // ��ȡ��һ�Ź��ʱ�
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

            // xlsx����������
            string SheetNames = Path.GetFileNameWithoutExtension(xlsxPath);
            UnityEngine.Debug.Log("��������:  " + SheetNames);

            FileStream stream = null;
            try
            {
                stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException e)
            {
                UnityEngine.Debug.LogErrorFormat("��ȡ�ļ�����:   " + e);
            }
            if (stream == null)
            {
                UnityEngine.Debug.LogErrorFormat("stream == null  ");
                return;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            // ��ȡ��һ�Ź��ʱ�
            Type t = Type.GetType(SheetNames);
            DataTable dataTable = result.Tables[0];

            string directoryPath = string.Format("{0}/{1}", UnityEngine.Application.streamingAssetsPath, "Jsons");
            string savePath = string.Format("{0}/{1}/{2}.json", UnityEngine.Application.streamingAssetsPath, "Jsons", SheetNames);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath); // ���Ŀ¼�����ڣ��򴴽�
                UnityEngine.Debug.Log("Ŀ¼�����ڣ��Ѵ�����" + directoryPath);
            }
            AssetDatabase.Refresh();

            ReadSingleSheet(t, dataTable, savePath );
            UnityEngine.Debug.Log("������� !��!  " + savePath);
            AssetDatabase.Refresh();


            excelReader.Close();
        }
    }



    /// <summary>
    /// ��ȡһ�������������
    /// </summary>
    /// <param name="type">Ҫת����struct��class����</param>
    /// <param name="dataTable">��ȡ�Ĺ���������</param>
    /// <param name="jsonPath">�洢·��</param>
    private static void ReadSingleSheet(Type type, DataTable dataTable, string jsonPath)
    {
        int rows = dataTable.Rows.Count;
        int Columns = dataTable.Columns.Count;
        // �������������
        DataRowCollection collect = dataTable.Rows;
        // xlsx��Ӧ�������ֶΣ��涨�ǵڶ���
        string[] jsonFileds = new string[Columns];
        // Ҫ�����Json��obj
        List<object> objsToSave = new List<object>();
        for (int i = 0; i < Columns; i++)
        {
            jsonFileds[i] = collect[1][i].ToString();
        }

        // �ӵ����п�ʼ
        for (int i = 3; i < rows; i++)
        {
            // ����һ��ʵ��
            Object objIns = type.Assembly.CreateInstance(type.ToString());

            for (int j = 0; j < Columns; j++)
            {
                // ��ȡ�ֶ�
                FieldInfo field = type.GetField(jsonFileds[j]);
                if (field != null)
                {
                    object value = null;
                    try // ��ֵ
                    {
                        UnityEngine.Debug.Log("��ֵ:  " + collect[i] + "  " + collect[i][j] + "   type: " + field.FieldType);
                        value = Convert.ChangeType(collect[i][j], field.FieldType);
                    }
                    catch (InvalidCastException e)
                    {
                        UnityEngine.Debug.LogWarningFormat("��ֵ����:  " + e.Message + "\n" + collect[i] + "  " + collect[i][j] + "   type: " + field.FieldType);

                        // ------ ����ͬ���͵����� ------
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
                            // ------ ���� bool[] ------
                            string str = collect[i][j].ToString();
                            string[] strs = str.Split(',');
                            bool[] bools = new bool[strs.Length];
                            for (int k = 0; k < strs.Length; k++)
                            {
                                bools[k] = bool.Parse(strs[k].Trim()); // ȥ������ո񣬲�����
                            }
                            value = bools;
                        }
                    }

                    field.SetValue(objIns, value);
                }
                else
                {
                    UnityEngine.Debug.LogFormat("���޷�ʶ����ַ�����{0}", jsonFileds[j]);
                }
            }
            objsToSave.Add(objIns);
        }
        // ����ΪJson
        string content = Newtonsoft.Json.JsonConvert.SerializeObject(objsToSave);
        SaveFile(content, jsonPath);
    }


    private static void GetFiles(string path, string suffix, ref List<string> fileNameList, bool isSubcatalog = false)
    {
        string filename;
        DirectoryInfo dir = new DirectoryInfo(path);
        FileInfo[] file = dir.GetFiles();
        //����������ļ���ʱ��Ҫʹ��  
        DirectoryInfo[] dii = dir.GetDirectories();

        foreach (FileInfo f in file)
        {
            filename = f.FullName;//�õ����ļ�������·��
            if (filename.EndsWith(suffix))//�ж��ļ���׺������ȡָ����ʽ���ļ�ȫ·��������fileList  
            {
                fileNameList.Add(filename);
                //UnityEngine. Debug.Log(filename);
            }
        }
        //��ȡ���ļ����ڵ��ļ��б��ݹ����
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
        // �������������
        DataRowCollection collect = dataTable.Rows;

        List<string> ziduan = new List<string>();
        List<string> name = new List<string>();
        List<string> annotation = new List<string>();
        for (int i = 0; i < Columns; i++)
        {
            UnityEngine.Debug.LogFormat("{3}��  {0}  ע��: {4}  �ֶ�: {2}   ����: {1} ", i, collect[1][i].ToString(), collect[2][i].ToString(), className , collect[0][i].ToString());
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
            //�滻����
            templateContent = templateContent.Replace("DataScriptTemplate", scriptName);
            templateContent = templateContent.Replace("}", " ");
            // ��ģ������������ֶ�
            string str = "";
            for (int i = 0; i < ziduan.Count; i++)
            {
                str += "    /// <summary>\n" + "    ///  " + annotation[i] + "\n    /// </summary>\n"; 
                str += "    public " + ziduan[i] + "  " + name[i] + " ; \n";
            }
            str += " \n  }";
            templateContent += str;



            string directoryPath = Path.GetDirectoryName(scriptPath); // ��ȡ�ļ�Ŀ¼·��
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath); // ���Ŀ¼�����ڣ��򴴽�
                UnityEngine.Debug.Log("Ŀ¼�����ڣ��Ѵ�����" + directoryPath);
            }
            //AssetDatabase.Refresh();

            try
            {
                // ������д���ļ�
                File.WriteAllText(scriptPath, templateContent, System.Text.Encoding.UTF8);
                UnityEngine.Debug.Log("�ļ��ɹ�д�룺" + scriptPath);
            }
            catch (Exception ex)
            {
                UnityEngine.Debug.LogError("д���ļ�ʧ�ܣ�" + ex.Message);
            }

            ////����֯�õ�����д���ļ�
            //File.WriteAllText(scriptPath, templateContent, System.Text.Encoding.UTF8);
            ////ˢ��һ����Դ����Ȼ�������ļ����һʱ�䲻����ʾ
            //AssetDatabase.Refresh();
        }
        else
        {
            UnityEngine.Debug.LogWarning("�ļ�������Ϊ��");
        }
    }

    #endregion 
}
#endif