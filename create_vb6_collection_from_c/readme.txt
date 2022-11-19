
I have some examples on how to work with collections in C

this has always entailed receiving a live collection either
as an argument, or by calling a vb6 function which returns a
collection as a return value.

I finally ended up getting around to working out how to have
the runtime itself create a live collection for me in C

this can be useful if you wanted to return a new collection object 
from a C function to VB. Of course you could always just pass in an 
empty collection as an argument to the function as well.

curosity killed the cat.

the previous example is here:
	http://sandsprite.com/blogs/index.php?uid=11&pid=388&year=2016

code: 
	http://sandsprite.com/blogs/files/col_dll.zip