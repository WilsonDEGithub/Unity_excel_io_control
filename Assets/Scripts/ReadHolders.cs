using UnityEngine;

/// <summary>
/// 读取booknames的scriptObject文件
/// 使用Resources直接读取
/// </summary>
public class ReadHolders : MonoBehaviour {
    readonly string assetName = "booknames";

	void Start ()
    {
        BookHolder asset = Resources.Load<BookHolder>(assetName);
        foreach (Menu gd in asset.menus)
        {
            Debug.Log(gd.m_Id);
            Debug.Log(gd.m_level);
            Debug.Log(gd.m_parentId);
            Debug.Log(gd.m_name);
        }
    }
}
