#ifndef CONVERT2ACCESS_H
#define CONVERT2ACCESS_H

#include "Config.h"

#include "Xlsx2DBTool.h"
#include "AccessSql.h"

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
class CConvert2Access {
public:
	CConvert2Access(CXlsx2DBTool *app);
	virtual ~CConvert2Access();

	int Convert(
		const std::string& sUrl,
		const std::string& sUserName,
		std::string& sPassword,
		std::string& sSource,
		config_item_t *item);

public:
	static void OutputCell(char *ptr, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol);
	static void OutputString(char *ptr, const char *str, field_meta_t& meta);
	static void OutputNumber(char *ptr, const double number, field_meta_t& meta);

public:
	CXlsx2DBTool *_app;
	CAccessSql *_accSql;

};

#endif
