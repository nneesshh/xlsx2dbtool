#ifndef XLSX2DBTOOL_H
#define XLSX2DBTOOL_H

#include "Config.h"

#include <string>
#include <vector>

enum CONVERT_2_TYPE {
	CONVERT_2_ACCESS = 1,
};

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
struct field_meta_t {
	int id;
	bool isField;
	bool isKey;
	bool isFromColumn;
	bool isToColumn;

	char name[256];
};

struct meta_table_t {
	struct field_meta_t arr[256];

	bool bTitleOk;
	int commentRowNum;
	int fromColumn;
	int toColumn;

};

struct config_item_t {
	CONVERT_2_TYPE c;

	char inputName[256];
	int inputSheetIdx;
	int inputTitleRowId;
	char inputFields[1024];
	char inputKeys[1024];
	char deleteClause[256];

	char sqlTableName[256];

	meta_table_t metaTable;
};

struct config_t {
	std::string _sUrl;
	std::string _sUserName;
	std::string _sPassword;
	std::string _sDataSource;

	std::vector<config_item_t> _vCfgItem;
};

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
class CXlsx2DBTool {
public:
	CXlsx2DBTool();
	virtual ~CXlsx2DBTool();

	void SetupMetaTable();

public:
	static void Usage(char *progName);
	
	static char * W2c(const wchar_t *wstr, char *cstr, int clenMax);
	static wchar_t * C2w(const char *cstr, wchar_t *wstr, int wlenMax);
	static int	UnicodeToGB2312(char* pOut, int nOut, const wchar_t *pIn, int nIn);
	static int	Gb2312ToUnicode(wchar_t* pOut, int nOut, const char *pIn, int nIn);

	static bool StartsWith(const TCHAR *str, const TCHAR *pattern);

public:
	config_t _config;
};

#endif // XLSX2DBTOOL_H
