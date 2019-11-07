using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;
using UnityEngine;

public class ReadExcel : MonoBehaviour {

	// Use this for initialization
	void Start ()
    {
        string url = Application.dataPath + "/ExcelToObject/Test.xlsx";

        int columnNum = 0, rowNum = 0;
        DataRowCollection collect = Read(url, ref columnNum, ref rowNum);
        List<Item> ItemList=new List<Item>();
        for (int i = 1; i < rowNum; i++)
        {
            Item item = new Item();
            //解析每列的数据
            item.Name = collect[i][0].ToString();
            item.Length = collect[i][1].ToString();
            item.Width = collect[i][2].ToString();
            ItemList.Add(item);
        }

        Debug.Log(ItemList.Find(x=>x.Name.Equals("给对方")).Length);
    }
    DataRowCollection Read(string filePath, ref int columnNum, ref int rowNum)
    {
        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        DataSet result = excelReader.AsDataSet();
        columnNum = result.Tables[0].Columns.Count;
        rowNum = result.Tables[0].Rows.Count;
        return result.Tables[0].Rows;
    }

    public class  Item
    {
        public string Name;

        public string Length;

        public string Width;

        public string Height;

        public string Id;
    }
}
