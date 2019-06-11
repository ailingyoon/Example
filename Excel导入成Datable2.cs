/// <summary>
/// DataTable写入Excel
/// </summary>
/// <param name="dt"></param>
/// <param name="strExcelFileName"></param>
/// <returns></returns>
public bool GridToExcelByNPOI(DataTable dt, string strExcelFileName)
{
    try
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Sheet1");

        ICellStyle HeadercellStyle = workbook.CreateCellStyle();
        HeadercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        HeadercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        HeadercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        HeadercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        HeadercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        //字体
        NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
        headerfont.Boldweight = (short)FontBoldWeight.Bold;
        HeadercellStyle.SetFont(headerfont);

        //用column name 作为列名
        int icolIndex = 0;
        IRow headerRow = sheet.CreateRow(0);
        foreach (DataColumn item in dt.Columns)
        {
            ICell cell = headerRow.CreateCell(icolIndex);
            cell.SetCellValue(item.ColumnName);
            cell.CellStyle = HeadercellStyle;
            icolIndex++;
        }

        ICellStyle cellStyle = workbook.CreateCellStyle();

        //为避免日期格式被Excel自动替换，所以设定 format 为 『@』 表示一率当成text來看
        cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
        cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

        NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
        cellfont.Boldweight = (short)FontBoldWeight.Normal;
        cellStyle.SetFont(cellfont);

        //建立内容行
        int iRowIndex = 1;
        int iCellIndex = 0;
        foreach (DataRow Rowitem in dt.Rows)
        {
            IRow DataRow = sheet.CreateRow(iRowIndex);
            foreach (DataColumn Colitem in dt.Columns)
            {

                ICell cell = DataRow.CreateCell(iCellIndex);
                cell.SetCellValue(Rowitem[Colitem].ToString());
                cell.CellStyle = cellStyle;
                iCellIndex++;
            }
            iCellIndex = 0;
            iRowIndex++;
        }

        //自适应列宽度
        for (int i = 0; i < icolIndex; i++)
        {
            sheet.AutoSizeColumn(i);
        }

        //写Excel
        FileStream file = new FileStream(strExcelFileName, FileMode.OpenOrCreate);
        workbook.Write(file);
        file.Flush();
        file.Close();
        return true;
    }
    catch (Exception ex)
    {
        return false;
    }
}


/// <summary>
/// 将excel中的数据导入到DataTable中
/// </summary>
/// <param name="fileName">fileName</param>
/// <param name="sheetName">excel工作薄sheet的名称</param>
/// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
/// <returns>返回的DataTable</returns>
public DataTable ExcelToDataTable(string fileName, string sheetName, bool isFirstRowColumn)
{
    ISheet sheet = null;
    DataTable data = new DataTable();
    IWorkbook workbook = null;
    FileStream fs = null;
    int startRow = 0;
    try
    {
        fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            workbook = new XSSFWorkbook(fs);
        else if (fileName.IndexOf(".xls") > 0) // 2003版本
            workbook = new HSSFWorkbook(fs);

        if (sheetName != null)
        {
            sheet = workbook.GetSheet(sheetName);
            if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
            {
                sheet = workbook.GetSheetAt(0);
            }
        }
        else
        {
            sheet = workbook.GetSheetAt(0);
        }
        if (sheet != null)
        {
            IRow firstRow = sheet.GetRow(0);
            int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

            if (isFirstRowColumn)
            {
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = cell.StringCellValue;
                        if (cellValue != null)
                        {
                            DataColumn column = new DataColumn(cellValue);
                            data.Columns.Add(column);
                        }
                    }
                }
                startRow = sheet.FirstRowNum + 1;
            }
            else
            {
                startRow = sheet.FirstRowNum;
            }

            //最后一列的标号
            int rowCount = sheet.LastRowNum;
            for (int i = startRow; i <= rowCount; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue; //没有数据的行默认是null　　　　　　　

                DataRow dataRow = data.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; ++j)
                {
                    if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        dataRow[j] = row.GetCell(j).ToString();
                }
                data.Rows.Add(dataRow);
            }
        }

        return data;
    }
    catch (Exception ex)
    {
        Console.WriteLine("Exception: " + ex.Message);
        return null;
    }
}






var Sql = "select id,name from test";
var dt = DB.DB.ExecuteDataTable(Sql);
//保存文件
string saveName = "temp_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
string DownloadExcelDic = "DownloadExcel";
string webDic = AppDomain.CurrentDomain.BaseDirectory + DownloadExcelDic;
if (!Directory.Exists(webDic))//判断目录是否存在
{
    Directory.CreateDirectory(webDic);//不存在则创建新目录
}
string path = webDic + @"\" + saveName; //目录+文件名+后缀名
bool isBuild = GridToExcelByNPOI(dt, path);




//1、先上传文件，将文件保存到服务器上

//2、导入Excel
string fileDic = "";//文件路径
var dt = ExcelToDataTable(fileDic, "sheet1", true);