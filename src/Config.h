#ifndef CONFIG_H
#define CONFIG_H

// config.h for MSVC.
#undef HAVE_ICONV
#undef HAVE_ASPRINTF

#ifdef _MSC_VER
#pragma warning (disable : 4005)
#pragma warning (disable : 4244)
#pragma warning (disable : 4996)

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#endif

#include <tchar.h>


#endif
