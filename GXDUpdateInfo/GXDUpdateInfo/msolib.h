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
