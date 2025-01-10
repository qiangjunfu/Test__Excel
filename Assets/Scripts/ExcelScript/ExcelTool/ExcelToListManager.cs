using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using UnityEngine;
using ExcelDataReader;
using Newtonsoft.Json;

public class ExcelToListManager : MonoBehaviour
{
    #region Mono单例

    private static ExcelToListManager instance;
    private static readonly object lockObj = new object();
    private static bool isApplicationQuitting = false; // 防止退出时创建新实例

    public static ExcelToListManager Instance
    {
        get
        {
            if (isApplicationQuitting)
            {
                Debug.LogWarning($"[ExcelToListManager] Instance requested after application quit. Returning null.");
                return null;
            }

            lock (lockObj)
            {
                if (instance == null)
                {
                    // 尝试从场景中找到实例
                    instance = FindObjectOfType<ExcelToListManager>();
                    if (instance == null)
                    {
                        // 如果未找到，自动创建
                        GameObject obj = new GameObject(nameof(ExcelToListManager));
                        instance = obj.AddComponent<ExcelToListManager>();
                        DontDestroyOnLoad(obj); // 保持跨场景存在
                    }
                }
                return instance;
            }
        }
    }

    private void Awake()
    {
        if (instance == null)
        {
            instance = this;
            DontDestroyOnLoad(this.gameObject); // 保持跨场景存在
            Initialize();
        }
        else if (instance != this)
        {
            Debug.LogWarning($"[ExcelToListManager] Duplicate instance detected. Destroying this object.");
            Destroy(gameObject);
        }
    }

    private void OnApplicationQuit()
    {
        isApplicationQuitting = true;
    }

    private void OnDestroy()
    {
        if (instance == this)
        {
            instance = null;
        }
    }

    /// <summary>
    /// 初始化逻辑
    /// </summary>
    private void Initialize()
    {
        // 在这里添加初始化逻辑
        Debug.Log("ExcelToListManager Initialized.");
    }
    #endregion

    [SerializeField, ReadOnly] private string folderPath = ""; // Excel 文件夹路径
    [SerializeField] List<TestData> m_TestDataList = new List<TestData>();

    void Start()
    {
        Init();
    }

    public void Init()
    {
        folderPath = Path.Combine(Application.streamingAssetsPath, "Excels");
        LoadAllExcelData();
    }


    private void LoadAllExcelData()
    {
        if (!Directory.Exists(folderPath))
        {
            Debug.LogError($"Excel 文件夹路径不存在：{folderPath}");
            return;
        }

        List<string> fileNameList = new List<string>();
        GetFiles(folderPath, "xlsx", ref fileNameList);

        foreach (string filePath in fileNameList)
        {
            string sheetName = Path.GetFileNameWithoutExtension(filePath);

            // 加载 Excel 数据
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException e)
            {
                Debug.LogError($"读取文件失败：{filePath} 错误信息：{e.Message}");
                continue;
            }

            if (stream == null)
            {
                Debug.LogError($"无法打开文件流：{filePath}");
                continue;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            if (result.Tables.Count == 0)
            {
                Debug.LogWarning($"Excel 文件 {filePath} 不包含工作表");
                continue;
            }

            // 读取第一张工作表
            DataTable dataTable = result.Tables[0];

            // 动态调用解析方法
            switch (sheetName)
            {
                case "TestData":
                    m_TestDataList = ParseDataTable<TestData>(dataTable);
                    Debug.Log($"成功加载 TestData，共 {m_TestDataList.Count} 条数据");
                    break;

                default:
                    Debug.LogWarning($"未定义的数据类型：{sheetName}");
                    break;
            }

            excelReader.Close();
        }
    }

    /// <summary>
    /// 解析 DataTable 为 List<T>
    /// </summary>
    /// <typeparam name="T">目标数据类型</typeparam>
    /// <param name="dataTable">DataTable 数据</param>
    /// <returns>解析后的数据列表</returns>
    private List<T> ParseDataTable<T>(DataTable dataTable) where T : class, new()
    {
        List<T> dataList = new List<T>();
        DataRowCollection rows = dataTable.Rows;

        if (rows.Count < 4) // 确保至少有四行：标题行、字段名行、字段类型行、数据行
        {
            Debug.LogWarning("工作表数据少于四行（第一行为标题，第二行为字段名，第三行为字段类型，第四行起为数据）");
            return dataList;
        }

        // 获取字段名（第二行）
        List<string> fieldNames = new List<string>();
        foreach (var item in rows[1].ItemArray) // 第二行为字段名
        {
            fieldNames.Add(item.ToString());
        }

        // 从第四行开始读取数据
        for (int i = 3; i < rows.Count; i++) // 第四行及之后为数据行
        {
            T instance = new T();
            DataRow row = rows[i];

            for (int j = 0; j < fieldNames.Count; j++)
            {
                string fieldName = fieldNames[j];
                if (string.IsNullOrEmpty(fieldName)) continue;

                FieldInfo field = typeof(T).GetField(fieldName);
                if (field != null)
                {
                    try
                    {
                        object value = row[j];
                        if (value == DBNull.Value || string.IsNullOrEmpty(value.ToString()))
                        {
                            // 如果单元格为空，使用字段默认值
                            value = field.FieldType.IsValueType ? Activator.CreateInstance(field.FieldType) : null;
                        }
                        else if (field.FieldType.IsArray)
                        {
                            // 处理数组类型字段
                            Type elementType = field.FieldType.GetElementType();
                            string[] stringValues = value.ToString().Split(','); // 按逗号分隔
                            Array array = Array.CreateInstance(elementType, stringValues.Length);
                            for (int k = 0; k < stringValues.Length; k++)
                            {
                                array.SetValue(Convert.ChangeType(stringValues[k].Trim(), elementType), k);
                            }
                            value = array;
                        }
                        else
                        {
                            // 转换数据类型
                            value = Convert.ChangeType(value, field.FieldType);
                        }

                        field.SetValue(instance, value);
                    }
                    catch (Exception e)
                    {
                        Debug.LogWarning($"字段赋值失败：{fieldName}，值：{row[j]}，错误信息：{e.Message}");
                    }
                }
                else
                {
                    Debug.LogWarning($"未找到匹配的字段：{fieldName} 在类型 {typeof(T).Name} 中");
                }
            }

            dataList.Add(instance);
        }

        return dataList;
    }




    /// <summary>
    /// 获取目录下的所有文件
    /// </summary>
    private void GetFiles(string path, string suffix, ref List<string> fileNameList)
    {
        string filename;
        DirectoryInfo dir = new DirectoryInfo(path);
        FileInfo[] file = dir.GetFiles();

        foreach (FileInfo f in file)
        {
            filename = f.FullName;
            if (filename.EndsWith(suffix))
            {
                fileNameList.Add(filename);
            }
        }
    }
}
