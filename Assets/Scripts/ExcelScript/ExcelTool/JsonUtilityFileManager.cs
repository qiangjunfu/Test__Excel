using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UnityEngine;

public class JsonUtilityFileManager : MonoBehaviour
{
    #region Mono单例

    private static JsonUtilityFileManager instance;
    private static readonly object lockObj = new object();
    private static bool isApplicationQuitting = false; // 防止退出时创建新实例

    public static JsonUtilityFileManager Instance
    {
        get
        {
            if (isApplicationQuitting)
            {
                Debug.LogWarning($"[JsonUtilityFileManager] Instance requested after application quit. Returning null.");
                return null;
            }

            lock (lockObj)
            {
                if (instance == null)
                {
                    // 尝试从场景中找到实例
                    instance = FindObjectOfType<JsonUtilityFileManager>();
                    if (instance == null)
                    {
                        // 如果未找到，自动创建
                        GameObject obj = new GameObject(nameof(JsonUtilityFileManager));
                        instance = obj.AddComponent<JsonUtilityFileManager>();
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
            Debug.LogWarning($"[JsonUtilityFileManager] Duplicate instance detected. Destroying this object.");
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

    private void Initialize()
    {
        Debug.Log("JsonUtilityFileManager Initialized.");

        Init();
    }
    #endregion

    [SerializeField] private string folderPath;
    [SerializeField] List<TestData> m_TestDataList = new List<TestData>();



    public void Init()
    {
        folderPath = Path.Combine(Application.streamingAssetsPath, "Jsons");
        LoadAllData();
    }

    public void LoadAllData()
    {
        DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
        FileInfo[] files = directoryInfo.GetFiles("*.json");

        foreach (FileInfo file in files)
        {
            string content = File.ReadAllText(file.FullName); // 读取 JSON 文件内容
            string dataType = Path.GetFileNameWithoutExtension(file.Name); // 使用文件名作为数据类型

            //Debug.Log($"--------- dataType: {dataType}  content: {content}");
            try
            {
                // 根据类型动态反序列化数据
                switch (dataType)
                {
                    case "TestData":
                        m_TestDataList.AddRange(JsonUtilityArray<TestData>(content));
                        Debug.Log($"加载 {m_TestDataList.Count} 个 {dataType} 条目。");
                        break;
                    default:
                        Debug.LogWarning($"Unsupported data type {dataType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.LogError($"Error processing file {file.Name}: {ex.Message}");
            }
        }
    }




    #region GetDataList

    public List<TestData> GetNetworkDataList()
    {
        return new List<TestData>(m_TestDataList);
    }

    #endregion




    #region Json

    [Serializable]
    private class Wrapper<T>
    {
        public List<T> data;
    }

    private List<T> JsonUtilityArray<T>(string json)
    {
        string wrappedJson = $"{{ \"data\": {json} }}"; // 包装成对象
        Wrapper<T> wrapper = JsonUtility.FromJson<Wrapper<T>>(wrappedJson);
        return wrapper.data;
    }
    public byte[] JsonToByteArray<T>(T message)
    {
        string jsonString = JsonUtility.ToJson(message);
        //print($"JsonToByteArray : {jsonString}");
        return Encoding.UTF8.GetBytes(jsonString);
    }
    public T ByteArrayToJson<T>(byte[] data)
    {
        string jsonString = Encoding.UTF8.GetString(data);
        //print($"ByteArrayToJson : {jsonString}");
        return JsonUtility.FromJson<T>(jsonString);
    }

    public void SaveDataToFile<T>(List<T> dataList, string filePath)
    {
        try
        {
            string directoryPath = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
                Debug.Log($"目录 {directoryPath} 已创建");
            }

            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close(); // 创建空文件
                Debug.Log($"文件 {filePath} 已创建");
            }


            StringBuilder jsonStringBuilder = new StringBuilder();
            jsonStringBuilder.Append("[\n");
            for (int i = 0; i < dataList.Count; i++)
            {
                string jsonString = JsonUtility.ToJson(dataList[i], true);
                //Debug.Log($"序列化 JSON: {jsonString}");

                // 如果不是最后一个元素，添加逗号
                if (i < dataList.Count - 1)
                {
                    jsonStringBuilder.Append(jsonString + ",\n");
                }
                else
                {
                    jsonStringBuilder.Append(jsonString + "\n");
                }
            }
            jsonStringBuilder.Append("]");
            string finalJson = jsonStringBuilder.ToString();

            File.WriteAllText(filePath, finalJson);

            Debug.Log($"数据已保存到 {filePath} \n{finalJson}");

        }
        catch (Exception ex)
        {
            Debug.LogError($"保存数据时发生错误：{ex.Message}");
        }
    }


    #endregion


}
