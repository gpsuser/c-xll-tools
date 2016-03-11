#define		kFunctionCount	1   
#define		kMaxFuncParms	13

const char	*gModuleName = "\024cTools Add-In";
const int	gFunctionCount = kFunctionCount,
			gMaxFuncParms = kMaxFuncParms;


// R		xloper	values, arrays and range references		
// B		double											
// I		signed short int								
// J		signed long int									
// #		function has macro sheet equivalence
// !		function is volatile



LPSTR functionParms[kFunctionCount][kMaxFuncParms] =
{
//	function title, argument types, function name, arg names, type (1=func,2=cmd),
//		group name (func wizard), hotkey, help ID, func help string, (repeat) argument help strings

	{" sort_ex",																	// function title
		" RRI#!",																	// argument types Return Type, First Parameter, Second Parameter    
		" sort_ex",																	// function name
		" X,t",																		// arg names
		" 1",																		// type (1=func,2=cmd)
    	" cTools",																	// group name (func wizard)
		" ",																		// hotkey
    	" cTools.hlp!315",															// help ID
    	" Sorts the first column of a two column array (float data).",				// func help string
    	" The two columns of float data.",											// arg 1 help string
		" sort Type: 1 = low to high & 2 = high to low."							// arg 2 help string
    },

};

LPSTR *gFuncParms = &functionParms[0][0];
