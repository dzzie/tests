
standard exe can place a form in ROT for external use with GetObject (thanks wqweto)
normally public class variables on this form can not be accesssed but controls can

individual classes from standard vb exe can not be accessed directly 

if we patch the class Object.ObjectType field to add bit 0x800 then this changes

any classes created after the patch  will have the new object type kick in.  

this sample will get to the object table from any convient internal object (here the Form itself)
and then load the whole object table, and patch every class it finds in the project.

This should be a handy way to allow remote scripting of your apps with minimal fuss and without having to compile
as an activeX exe which I dislike. long term stability and possible weird nuances not yet tested 




