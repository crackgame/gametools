#include <afxdb.h> 
#include <odbcinst.h>


//CDatabase database;

CRecordset* xlsOpen(const char* xlsFile, const char* sheetName);
void xlsclose();


//void xlsGetFiledValue(short nIndex,CString &str);
//void xlsMoveNext();

