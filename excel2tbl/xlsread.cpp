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
    
    // �������д�ȡ���ַ���
	strConnection="Driver={Microsoft Excel Driver (*.xls)};ReadOnly=0;DBQ=";
	strConnection+=xlsFile;
    
	TRY
    {
	
        // �����ݿ�(��Excel�ļ�)
        database.Open(NULL, false, false, strConnection);
	
        // ���ö�ȡ�Ĳ�ѯ���.
		sSql.Format("SELECT * FROM [%s$]", sheetName);

        // ִ�в�ѯ���
        recset.Open(CRecordset::forwardOnly, sSql, CRecordset::readOnly);
	
		return &recset;
	}
    CATCH(CDBException, e)
    {
        // ���ݿ���������쳣ʱ...
        AfxMessageBox("���ݿ����: " + e->m_strError,MB_ICONERROR|MB_OK);
		return NULL;
    }
    END_CATCH;

	return &recset;
}