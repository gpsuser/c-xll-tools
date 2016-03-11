/*
	NOTE:
	=====
	
	Parts of this project rely on and uses code that was produced by the following sources:
	
	1. John Champion:	http://www.codeproject.com/Articles/5263/Sample-Excel-Function-Add-in-in-C
	2. Microsoft Excel97 Developer's Kit

	The authors of this project would like to thank these two sources for their contribution. 
*/

#include <stdlib.h> //refernce qsort() here

#include "cTools.h"
#include "intrface.h"

// #define FREE_ARG char*
//#define NR_END 1

HANDLE	hArray2;
HANDLE	hArray;				// temporary handle for returning large arrays to XL

// ========================================================================================
//
// This function is called by Excel if the xlbitDLLFree bit has been set in the
// return array of the interp function.  It allows us to free up allocated memory.
//
// ========================================================================================
__declspec(dllexport) void xlAutoFree( LPXLOPER pxFree )
{
	if ( hArray )
	{
		GlobalUnlock( hArray );
		GlobalFree( hArray );
		hArray = 0;
	}
	return;
}

// ========================================================================================
//
// ClipSize is a utility function that will determine the size of a "multi" array
// structure.  It checks to see if the data is organized in columns or rows (giving
// preference to columns), and ignores empty cells at the end of the array.
// It returns the size of the 1D table of valid data.
//
// ========================================================================================

WORD ClipSize( XLOPER *multi )
{
	WORD		size, i;
	LPXLOPER	ptr;
	
	// get the number of columns in the data
	size = multi->val.array.columns;
	
	// if there's only one column, then it must be organized in multiple rows.
	if ( size == 1 )
		size = multi->val.array.rows;

	// ignore empty or error values at the end of the array.
	for ( i = size - 1; i >= 0; i-- )
	{
		ptr = multi->val.array.lparray + i;

		if ( ptr->xltype != xltypeNil )
			break;
	}
	
	return i + 1;
}
/*
The sorting function that returns a sorted array of floats, given an unsorted two column array
The sorting is done by qsort() in stdlib.h

parameter: x			The two column array of floating point data
parameter: sortType		1: low to high sorting, 2: high to low sorting
return	 :				The sorted two column array	

*/
__declspec(dllexport) LPXLOPER sort_ex(LPXLOPER x, int sortType){

	static XLOPER	resultBuffer[dimResult],		// Return Data Array
					xMulti,							// x coerced to xltypeMulti
					tempTypeMulti,					// xltypeMulti in an XLOPER for passing
					retMulti;						// return data structure
	


	short			hasxMulti = 0,					// flags to indicate memory has been allocated
					error = -1;						// -1 if no error; error code otherwise
    	
	LPXLOPER		retArrayPtr, 					// pointer to the return results array
					xPtr;							// pointer to array of input x data

	ULONG			numColumns_x, numRows_x,		// number of columns and rows of the data matrix received from excel
					Count = 0;						// the dimension of the results array 

	
	Data* s;										// the data structure that holds the column data
	ULONG i;
	LPXLOPER tempPtr;


	// Initialize variables
	tempTypeMulti.xltype = xltypeInt;
    tempTypeMulti.val.w = xltypeMulti;
	
	// Make sure x data is an expected data type
	if (	x->xltype != xltypeRef	&&	x->xltype != xltypeSRef	&&	x->xltype != xltypeMulti && x->xltype != xltypeNum ){
		error = xlerrValue;
		goto done;
	}
        
	// Convert the x data to a "multi" type & indicate Excel has allocated memory for the targMulti
	if ( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &xMulti, 2,(LPXLOPER) x, (LPXLOPER) &tempTypeMulti ) ) hasxMulti = 1;	

	// save a temporary pointer to the x data values
	xPtr = xMulti.val.array.lparray;
	
	// number of rows and columns in the excel data
	numRows_x=xMulti.val.array.rows;
	numColumns_x=xMulti.val.array.columns;

	// malloc memory for the data object
	s = (Data *) malloc (numRows_x*sizeof(Data));
 
	// must have the correct dimensions for solution arrray when considering (r1xc1) X (r2xc2)=r1xc2
	Count =  numColumns_x*numRows_x ;
		
	// set up the return array data structure, type is "multi"
	retMulti.xltype = xltypeMulti;
	
	// row and column size of return data must match those of the multiplied matrix								
	retMulti.val.array.rows = (int)numRows_x;
	retMulti.val.array.columns = (int)numColumns_x;
		    
	// For efficiency, we have a static buffer that holds up to "dimResult" values. 
	// If it's large enough, use it.  Otherwise, allocate memory and tell excel to call us back to free it ( via xlAutoFree ).
	if (  Count > dimResult ){
		retArrayPtr = (LPXLOPER) GlobalLock( hArray2 =GlobalAlloc( GMEM_ZEROINIT,  Count * sizeof(XLOPER)) );
		retMulti.xltype |= xlbitDLLFree;
	}else{
		hArray2 = 0;
		retArrayPtr = resultBuffer;
	}
			                                                           
	retMulti.val.array.lparray = retArrayPtr;

	// initialise the DATA object: TWO COLS OF FLOAT DATA
	for(i=0;i<numRows_x;i++){													
		 s[i].f1 =(float)xPtr[i*numColumns_x + 0].val.num;
		 s[i].f2 =(float)xPtr[i*numColumns_x + 1].val.num;
		 s[i].compareLtoH =lowToHighCompareFn;
		 s[i].compareHtoL =highToLowCompareFn;
	}

	// in place sort
	if(sortType == 1) { //sort low to high
		qsort(s,numRows_x,sizeof(Data),s[0].compareLtoH);
	} else{//sort high to low
		qsort(s,numRows_x,sizeof(Data),s[0].compareHtoL);

	}

	for(i=0;i<numRows_x;i++){						 // rows
		tempPtr = &retArrayPtr[i*numColumns_x + 0];	 // tempPtr points to the current element of the return data "multi" structure
		tempPtr->xltype = xltypeNum;				 // the data type will be ordinary numeric (floating point) data
		tempPtr->val.num = s[i].f1;					 // STORE VALUE TO RETURN ARRAY

		tempPtr = &retArrayPtr[i*numColumns_x + 1];	 // tempPtr points to the current element of the return data "multi" structure
		tempPtr->xltype = xltypeNum;				 // the data type will be ordinary numeric (floating point) data
		tempPtr->val.num = s[i].f2;
	}

done:
    
	// free the memory allocated by Excel on our behalf
	if ( hasxMulti )	Excel4( xlFree, 0, 1, (LPXLOPER) &xMulti );

	 
 free(s);
	
	// if the "error" variable was set above, something significant failed and we should return an error for all x targets
	if ( error != -1 ){
			resultBuffer->xltype = xltypeErr;
			resultBuffer->val.err = error;
		return (LPXLOPER) resultBuffer;
    }

	// if there was more than one x (target) value, return a multi structure, otherwise just return a single XLOPER
	if ( Count > 1 )
		return (LPXLOPER) &retMulti;
	else
		return (LPXLOPER) resultBuffer;
}


// compare function to enable sorting from low to high
int lowToHighCompareFn(const void *d1, const void *d2){

	Data * D1 = (Data *) d1;		// casts the void object to type Data
	Data * D2 = (Data *) d2;

	if(D1->f1 > D2->f1)
		return 1;
	else if(D1->f1 < D2->f1)
		return -1;

	return 0;
}

//  compare function to enable sorting from high to low
int highToLowCompareFn(const void *d1, const void *d2){

	Data * D1 = (Data *) d1;		// casts the void object to type Data
	Data * D2 = (Data *) d2;

	if(D1->f1 < D2->f1)
		return 1;
	else if(D1->f1 > D2->f1)
		return -1;

	return 0;
}
