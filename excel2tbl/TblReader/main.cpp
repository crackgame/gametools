// TblReader.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "DataManager.h"

#pragma pack(1)
struct stItemBase 
{
	DWORD id;
	char name[50];
	char detail[255];
};
#pragma pack()

int _tmain(int argc, _TCHAR* argv[])
{
	DataManager<stItemBase> itemManager;
	if ( !itemManager.load("./tables/item.tbl") ) {
		printf("读取物品表失败\n");
		getchar();
		return 0;
	}

	stItemBase* pItemB = itemManager.getData(1);
	if (pItemB != NULL) {
		printf("id=%u, name=%s, detail=%s\n", pItemB->id, pItemB->name, pItemB->detail);
	}

	pItemB = itemManager.getData(2);
	if (pItemB != NULL) {
		printf("id=%u, name=%s, detail=%s\n", pItemB->id, pItemB->name, pItemB->detail);
	}

	pItemB = itemManager.getData(3);
	if (pItemB != NULL) {
		printf("id=%u, name=%s, detail=%s\n", pItemB->id, pItemB->name, pItemB->detail);
	}

	getchar();
	return 0;
}

