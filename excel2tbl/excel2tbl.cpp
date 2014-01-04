// excel2tbl.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "xlsread.h"
#include <string>
#include <vector>
#include <tinyxml.h>
#include <direct.h>

using namespace std;

#define TEMP_PATH	".\\temp\\"


std::string AnsiToUtf8(std::string strAnsi)
{
	std::string ret;
	if (strAnsi.length() > 0)
	{	
		int nWideStrLength = MultiByteToWideChar(CP_ACP, 0, strAnsi.c_str(), -1, NULL, 0);
		WCHAR* pwszBuf = (WCHAR*)malloc((nWideStrLength+1)*sizeof(WCHAR));
		memset(pwszBuf, 0, (nWideStrLength+1)*sizeof(WCHAR));
		MultiByteToWideChar(CP_ACP, 0, strAnsi.c_str(), -1, pwszBuf, (nWideStrLength+1)*sizeof(WCHAR));

		int nUtf8Length = WideCharToMultiByte( CP_UTF8,0,pwszBuf,-1,NULL,0,NULL,FALSE );
		char* pszUtf8Buf = (char*)malloc((nUtf8Length+1)*sizeof(char));
		memset(pszUtf8Buf, 0, (nUtf8Length+1)*sizeof(char));

		WideCharToMultiByte(CP_UTF8, 0, pwszBuf, -1, pszUtf8Buf, (nUtf8Length+1)*sizeof(char), NULL, FALSE);
		ret = pszUtf8Buf;

		free(pszUtf8Buf);
		free(pwszBuf);
	}
	return ret;
}


void DeleteDirFile(CString sPath)
{
	WIN32_FIND_DATA fd;
	HANDLE hFind = ::FindFirstFile(sPath + "*.*",&fd);

	if (hFind != INVALID_HANDLE_VALUE)
	{           
		while (::FindNextFile(hFind,&fd))
		{
			//�ж��Ƿ�ΪĿ¼
			if (fd.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY) {
				CString name;
				name = fd.cFileName;
				//�ж��Ƿ�Ϊ.��..
				if ((name != ".") && (name != ".."))
				{
					//�����������Ŀ¼�����еݹ�
					DeleteDirFile(sPath + fd.cFileName + "\\");
				}
			} else {
				DeleteFile(sPath + fd.cFileName);
			}
		}
		::FindClose(hFind);
	}
	RemoveDirectory(sPath);
}


struct stField 
{
	string name;
	string type;
	UINT size;

	stField()
	{
		name = "";
		type = "";
		size = 0;
	}
};

int _tmain(int argc, _TCHAR* argv[])
{
	//// �������б�.xls client_table.xml
	//argc = 2;
	//argv[0] = "�������б�.xls";
	//argv[1] = "client_table.xml";

	if (argc != 2) {
		printf("��������ȷ\n");
		printf("����: excel2tbl.exe server_maker.xml \n");
		return 0;
	}

	string strXml = argv[1];
	string strOutPath = "./";
	strOutPath += strXml.substr(0,strXml.length()-4);
	strOutPath += "/";

	DeleteDirFile(strOutPath.c_str());
	_mkdir(strOutPath.c_str());

	TiXmlDocument doc;
	if ( !doc.LoadFile(strXml.c_str()) ) {
		printf("��ȡ�����ļ�ʧ�� %s\n", strXml.c_str());
		return 0;
	}

	TiXmlElement* root = doc.FirstChildElement("root");
	if (root == NULL) {
		printf("û�и��ڵ� root");
		return 0;
	}

	// ������ʱ�ļ���
	DeleteDirFile(TEMP_PATH);
	_mkdir(TEMP_PATH);

	CRecordset* rs = NULL;
	TiXmlNode* node = NULL;
	TiXmlElement* database = NULL;
	TiXmlElement* table = NULL;
	TiXmlElement* field = NULL;
	string strXls, strSheet, strOut;
	database = (TiXmlElement*)root->FirstChild("database");
	while( database ) {
		if (database->Type() == TiXmlNode::TINYXML_COMMENT) {
			database = (TiXmlElement*)database->NextSibling();
			continue;
		}

		strXls = database->Attribute("xls");

		table = (TiXmlElement*)database->FirstChild("table");
		while( table ) {
			if (table->Type() == TiXmlNode::TINYXML_COMMENT) {
				table = (TiXmlElement*)table->NextSibling();
				continue;
			}

			strSheet = table->Attribute("sheet");
			strOut = strOutPath;
			strOut += table->Attribute("out");

			TRY
			{
				// Ϊ�˷�ֹ���Ѵ��ļ�ʱ����ռ���棬����һ�ݳ������д���
				char tempFile[MAX_PATH];
				sprintf(tempFile, "%s%s", TEMP_PATH, strXls.c_str());
				CopyFile(strXls.c_str(), tempFile, FALSE);

				rs = xlsOpen(tempFile, strSheet.c_str());
				if (rs == NULL) {
					printf("�򿪱��ʧ�ܣ���û�й����� xls=%s, sheet=%s\n", strXls.c_str(), strSheet.c_str());
					return 0;
				}
			}
			CATCH(CDBException, e)
			{
				// ���ݿ���������쳣ʱ...
				CString strError;
				strError.Format("���ݿ����: %s", e->m_strError);
				AfxMessageBox(strError,MB_ICONERROR|MB_OK);
				return 0;
			}
			END_CATCH;

			FILE* pFile = fopen(strOut.c_str(), "w+b");
			if (pFile == NULL) {
				printf("�����ļ�ʧ�� file=%s", strOut.c_str());
				return 0;
			}

			// ������ʾ
			printf("���ڵ��� excel=%s, sheet=%s, tbl=%s\n",
				strXls.c_str(), strSheet.c_str(), strOut.c_str());

			vector<stField> vecField;

			// ȡ���ֶ�
			field = (TiXmlElement*)table->FirstChild("field");
			while( field ) {
				if (field->Type() == TiXmlNode::TINYXML_COMMENT) {
					field = (TiXmlElement*)field->NextSibling();
					continue;
				}

				stField sField;
				sField.name  = field->Attribute("name");
				sField.type = field->Attribute("type");
				sField.size = (UINT)atoi(field->Attribute("len"));
				vecField.push_back(sField);
				field = (TiXmlElement*)field->NextSibling();
			}

			if (vecField.empty()) {
				fclose(pFile);
				printf("û���ֶ���Ҫ����\n");
				return 0;
			}

			DWORD count = 0;
			fwrite(&count, sizeof(count), 1, pFile);

			UINT curID = 0;
			string curFieldName = "";

			TRY
			{
				while(!rs->IsEOF())
				{
					for (size_t i=0; i<vecField.size(); i++) {
						stField sFiled = vecField[i];
						curID = (UINT)i;
						curFieldName = sFiled.name;
						CString strValue;
						rs->GetFieldValue(sFiled.name.c_str(),strValue);
						if (sFiled.type == "INT") {
							UINT num = strtoul(strValue.GetBuffer(0), NULL, 10);
							fwrite(&num, sizeof(num), 1, pFile);
						} else if (sFiled.type == "FLOAT") {
							FLOAT num = (FLOAT)strtod(strValue.GetBuffer(0), NULL);
							fwrite(&num, sizeof(num), 1, pFile);
						} else if (sFiled.type == "BIGINT") {
							UINT64 num = 0;
							sscanf(strValue.GetBuffer(0), "%llu", &num);
							//UINT64 num = (UINT64)strtod(strValue.GetBuffer(0), NULL);
							fwrite(&num, sizeof(num), 1, pFile);
						} else if (sFiled.type == "ANSI") {
							char* tmp = new char[sFiled.size+1];
							strncpy(tmp, strValue.GetBuffer(0), sFiled.size);
							fwrite(tmp, sFiled.size, 1, pFile);
							delete[] tmp;
						} else if (sFiled.type == "UTF8") {
							string strUTF8 = AnsiToUtf8(strValue.GetBuffer(0));
							char* tmp = new char[sFiled.size+1];
							strncpy(tmp, strUTF8.c_str(), sFiled.size);
							fwrite(tmp, sFiled.size, 1, pFile);
							delete[] tmp;
						} else {
							printf("��Ч���ֶ����� type=%s\n", sFiled.type.c_str());
						}
					}

					count++;
					rs->MoveNext();
				}

				rs->Close();
				xlsclose();
			}
			CATCH(CDBException, e)
			{
				fclose(pFile);

				printf("����� %u �����ݵ� \"%s\" �ֶ�ʱ������\n", curID, curFieldName.c_str());

				// ���ݿ���������쳣ʱ...
				CString strError;
				strError.Format("��������ʱ����: ���ݿ����: %s", e->m_strError);
				AfxMessageBox(strError,MB_ICONERROR|MB_OK);
				return 0;
			}
			END_CATCH;

			fseek(pFile, 0, SEEK_SET);
			fwrite(&count, sizeof(count), 1, pFile);
			fseek(pFile, 0, SEEK_END);

			// �ر��ļ�
			fclose(pFile);

			table = (TiXmlElement*)table->NextSibling();
		}
		database = (TiXmlElement*)database->NextSibling();
	}

	// ɾ����ʱ�ļ���
	DeleteDirFile(TEMP_PATH);

	printf("�������ݵ����ɹ���\n");
#ifdef _DEBUG
	getchar();
#endif
	return 0;
}

