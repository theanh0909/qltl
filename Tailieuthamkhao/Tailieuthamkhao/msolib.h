/*
#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE12\MSO.DLL" \
	rename("DocumentProperties","DocumentPropertiesXL") \
	rename("RGB","RBGXL")
#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB"
#import "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE" \
	rename("DialogBox","DialogBoxXL") \
	rename("RGB","RBGXL") \
	rename("DocumentProperties","DocumentPropertiesXL") \
	rename("ReplaceText","ReplaceTextXL") \
	rename("CopyFile","CopyFileXL") \
	exclude("IFont","IPicture") \
	no_dual_interfaces

using namespace Excel;
*/

#import "..\\Office\\MSO.DLL" \
	rename("DocumentProperties","DocumentPropertiesXL") \
	rename("RGB","RBGXL")
#import "..\\Office\\VBE6EXT.OLB"
#import "..\\Office\\EXCEL.EXE" \
	rename("DialogBox","DialogBoxXL") \
	rename("RGB","RBGXL") \
	rename("DocumentProperties","DocumentPropertiesXL") \
	rename("ReplaceText","ReplaceTextXL") \
	rename("CopyFile","CopyFileXL") \
	exclude("IFont","IPicture") \
	no_dual_interfaces

using namespace Excel;
