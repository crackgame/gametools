#include "stdafx.h"
#include "xlsread.h"

CDatabase database;
CRecordset recset(&database);

//void xlsGetFiledValue(short nIndex,CString &str)
//{
//	recset.GetFieldValue(nIndex,str);
//}
//void xlsMoveNext()
//{
//	recset.MoveNext();
//}

void xlsclose()
{
      database.Close();

}

CRecordset* xlsOpen(const char* xlsFile, const char* sheetName)
{
    CString sSql;
    CString strConnection;
    
    // 创建进行存取的字符串
	strConnection="Driver={Microsoft Excel Driver (*.xls)};ReadOnly=0;DBQ=";
	strConnection+=xlsFile;
    
	TRY
    {
	
        // 打开数据库(既Excel文件)
        database.Open(NULL, false, false, strConnection);
	
        // 设置读取的查询语句.
		sSql.Format("SELECT * FROM [%s$]", sheetName);

        // 执行查询语句
        recset.Open(CRecordset::forwardOnly, sSql, CRecordset::readOnly);
	
		return &recset;
	}
    CATCH(CDBException, e)
    {
        // 数据库操作产生异常时...
        AfxMessageBox("数据库错误: " + e->m_strError,MB_ICONERROR|MB_OK);
		return NULL;
    }
    END_CATCH;

	return &recset;
}