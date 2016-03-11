#include "generic.h"

#define dimResult	200		// size of the permanent array for returning data to XL

WORD	ClipSize( XLOPER *multi );

typedef struct data{		
	float f1;
	float f2;
	int(*compareLtoH)(const void*, const void*);
	int(*compareHtoL)(const void*, const void*);
} Data;

int lowToHighCompareFn(const void *s1, const void *s2);
int highToLowCompareFn(const void *s1, const void *s2);

__declspec(dllexport) void xlAutoFree( LPXLOPER pxFree );
__declspec(dllexport) LPXLOPER sort_ex(LPXLOPER x, int sortType);

