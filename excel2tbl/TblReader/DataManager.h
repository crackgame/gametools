#pragma once
#include <map>
using namespace std;

// SData 必须包含id字段

template<typename SData>
class DataManager
{
public:
	DataManager(void){}
	virtual ~DataManager(void){m_dataMap.clear();}

	// 读取数据表
	bool load(const char* tblName)
	{
		unsigned int count = 0;
		FILE* pFile = fopen(tblName, "r+b");
		if (pFile == NULL)
			return false;

		fread(&count,sizeof(count),1,pFile);

		if (count > 0) 
			m_dataMap.clear();

		SData* datas = new SData[count];
		fread(datas, sizeof(SData)*count, 1, pFile);
		fclose(pFile);

		for (unsigned int i=0; i<count; i++) {
			m_dataMap[datas[i].id] = datas[i];
		}

		delete[] datas;

		return true;
	}

	// 取得指定ID的数据
	SData* getData(DWORD id)
	{
		map<DWORD,SData>::iterator itFind = m_dataMap.find(id);
		if (itFind == m_dataMap.end())
			return NULL;
        return &itFind->second;
	}

	DWORD size()
	{
		return (DWORD)m_dataMap.size();
	}

private:
	map<DWORD,SData> m_dataMap;
};
