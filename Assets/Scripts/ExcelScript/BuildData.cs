using System.Collections;
using System.Collections.Generic;
using UnityEngine;


[System.Serializable]
public class BuildData
{  
    /// <summary>
    ///  索引
    /// </summary>
    public int  id ; 
    /// <summary>
    ///  名字
    /// </summary>
    public string  name ; 
    /// <summary>
    ///  描述
    /// </summary>
    public string  describe ; 
    /// <summary>
    ///  数据类型
    /// </summary>
    public int  dataType ; 
    /// <summary>
    ///  对应的icon图标
    /// </summary>
    public string  iconPath ; 
    /// <summary>
    ///  物体路径
    /// </summary>
    public string  objPath ; 
    /// <summary>
    ///  建筑结构路径(隐藏结构)
    /// </summary>
    public string  floorStructPath ; 
    /// <summary>
    ///  3d光标路径
    /// </summary>
    public string  guangBiaoPath ; 
    /// <summary>
    ///  (分层)物体路径
    /// </summary>
    public string  objPath_FenCeng ; 
    /// <summary>
    ///  (分层)相机目标点路径
    /// </summary>
    public string  cameraTargetPath_FenCeng ; 
    /// <summary>
    ///  子楼层名字
    /// </summary>
    public string[]  subFloorName ; 
    /// <summary>
    ///  子楼层路径
    /// </summary>
    public string[]  subFloorPath ; 
    /// <summary>
    ///  楼层对应id (取10000开始, 千位数:建筑楼栋id , 百位数:楼层id)
    /// </summary>
    public int[]  subFloorId ; 
    /// <summary>
    ///  内部配套设施Id(
    /// </summary>
    public int[]  neiBuPeiTaoIdList ; 
    /// <summary>
    ///  测试数据1
    /// </summary>
    public float[]  test1 ; 
    /// <summary>
    ///  测试数据2
    /// </summary>
    public float[]  test2 ; 
    /// <summary>
    ///  测试数据3
    /// </summary>
    public float  test3 ; 
 
  }