Description of the **sort\_ex()** function.

  * This function sorts two columns of floating point data in a spreadsheet.

  * The sorting is always done on the first column of the data.

  * The _sort type_ parameter determines whether or not the sorting is done from high to low or from low to high.

  * This function relies on _qsort()_ in _stdlib.h_, to do the heavy lifting


_function signature_: **sort\_ex(X,t)**

  * **X**: The two columns of float data
  * **t** : The sort choice (1=low to high, 2= hight to low)
