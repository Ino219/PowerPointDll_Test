#pragma once

#ifdef POWERPOINTLIB_EXPORTS
#define POWERPOINTLIB_API __declspec(dllexport)
#else
#define POWERPOINTLIB_API __declspec(dllimport)
#endif



extern "C" __declspec(dllexport) int init(int a, int b) {
	return a + b;
}
;
