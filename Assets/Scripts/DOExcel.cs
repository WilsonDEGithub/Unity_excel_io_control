using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.Data;
using System.IO;
using Excel;

public class DoExcel
{
    public static DataSet ReadExcel(string path)
    {
        FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
        Debug.LogError("stream    " + stream);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);          //读取2007以后版本
                                                                                                //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);     //读取2003以后版本

        DataSet result = excelReader.AsDataSet();
        Debug.LogError("result   " + result);
        excelReader.Close();


        int columns = result.Tables[0].Columns.Count;
        int rows = result.Tables[0].Rows.Count;

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                string nvalue = result.Tables[0].Rows[i][j].ToString();
                Debug.Log(nvalue);
            }
        }
        return result;
    }

    public static List<DepenceTableData> Load(string path)
    {
        Debug.LogError(path);
        List<DepenceTableData> _data = new List<DepenceTableData>();
        DataSet resultds = ReadExcel(path);
        int column = resultds.Tables[0].Columns.Count;
        int row = resultds.Tables[0].Rows.Count;
        Debug.LogWarning(column + "  " + row);
        for (int i = 1; i < row; i++)
        {
            DepenceTableData temp_data;
            temp_data.instruct = resultds.Tables[0].Rows[i][0].ToString();
            temp_data.word = resultds.Tables[0].Rows[i][1].ToString();
            Debug.Log(temp_data.instruct + "  " + temp_data.word);
            _data.Add(temp_data);
        }
        return _data;
    }

}

public struct DepenceTableData
{
    public string word;
    public string instruct;
}