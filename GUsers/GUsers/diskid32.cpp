#include "pch.h"
#include "diskid32.h"
#include "HardDisk.h"

#define  IDENTIFY_BUFFER_SIZE	512

char HardDriveSerialNumber [MAX_LEN];
char HardDriveModelNumber [MAX_LEN];

//  Required to ensure correct PhysicalDrive IOCTL structure setup
#pragma pack(1)

//  IOCTL commands
#define  DFP_GET_VERSION          0x00074080
#define  DFP_SEND_DRIVE_COMMAND   0x0007c084
#define  DFP_RECEIVE_DRIVE_DATA   0x0007c088

#define  FILE_DEVICE_SCSI              0x0000001b
#define  IOCTL_SCSI_MINIPORT_IDENTIFY  ((FILE_DEVICE_SCSI << 16) + 0x0501)
#define  IOCTL_SCSI_MINIPORT 0x0004D008  //  see NTDDSCSI.H for definition

#define SMART_GET_VERSION               CTL_CODE(IOCTL_DISK_BASE, 0x0020, METHOD_BUFFERED, FILE_READ_ACCESS)
#define SMART_SEND_DRIVE_COMMAND        CTL_CODE(IOCTL_DISK_BASE, 0x0021, METHOD_BUFFERED, FILE_READ_ACCESS | FILE_WRITE_ACCESS)
#define SMART_RCV_DRIVE_DATA            CTL_CODE(IOCTL_DISK_BASE, 0x0022, METHOD_BUFFERED, FILE_READ_ACCESS | FILE_WRITE_ACCESS)

//  GETVERSIONOUTPARAMS contains the data returned from the 
//  Get Driver Version function.
typedef struct _GETVERSIONOUTPARAMS
{
   BYTE bVersion;      // Binary driver version.
   BYTE bRevision;     // Binary driver revision.
   BYTE bReserved;     // Not used.
   BYTE bIDEDeviceMap; // Bit map of IDE devices.
   DWORD fCapabilities; // Bit mask of driver capabilities.
   DWORD dwReserved[4]; // For future use.
} GETVERSIONOUTPARAMS, *PGETVERSIONOUTPARAMS, *LPGETVERSIONOUTPARAMS;


//  Bits returned in the fCapabilities member of GETVERSIONOUTPARAMS 
#define  CAP_IDE_ID_FUNCTION             1  // ATA ID command supported
#define  CAP_IDE_ATAPI_ID                2  // ATAPI ID command supported
#define  CAP_IDE_EXECUTE_SMART_FUNCTION  4  // SMART commannds supported

//  Valid values for the bCommandReg member of IDEREGS.
#define  IDE_ATAPI_IDENTIFY  0xA1  //  Returns ID sector for ATAPI.
#define  IDE_ATA_IDENTIFY    0xEC  //  Returns ID sector for ATA.

// The following struct defines the interesting part of the IDENTIFY
// buffer:
typedef struct _IDSECTOR
{
   USHORT  wGenConfig;
   USHORT  wNumCyls;
   USHORT  wReserved;
   USHORT  wNumHeads;
   USHORT  wBytesPerTrack;
   USHORT  wBytesPerSector;
   USHORT  wSectorsPerTrack;
   USHORT  wVendorUnique[3];
   CHAR    sSerialNumber[20];
   USHORT  wBufferType;
   USHORT  wBufferSize;
   USHORT  wECCSize;
   CHAR    sFirmwareRev[8];
   CHAR    sModelNumber[40];
   USHORT  wMoreVendorUnique;
   USHORT  wDoubleWordIO;
   USHORT  wCapabilities;
   USHORT  wReserved1;
   USHORT  wPIOTiming;
   USHORT  wDMATiming;
   USHORT  wBS;
   USHORT  wNumCurrentCyls;
   USHORT  wNumCurrentHeads;
   USHORT  wNumCurrentSectorsPerTrack;
   ULONG   ulCurrentSectorCapacity;
   USHORT  wMultSectorStuff;
   ULONG   ulTotalAddressableSectors;
   USHORT  wSingleWordDMA;
   USHORT  wMultiWordDMA;
   BYTE    bReserved[128];
} IDSECTOR, *PIDSECTOR;


typedef struct _SRB_IO_CONTROL
{
   ULONG HeaderLength;
   UCHAR Signature[8];
   ULONG Timeout;
   ULONG ControlCode;
   ULONG ReturnCode;
   ULONG Length;
} SRB_IO_CONTROL, *PSRB_IO_CONTROL;


// Define global buffers.
BYTE IdOutCmd [sizeof (SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE - 1];

char *ConvertToString (DWORD diskdata [256],
			   int firstIndex,
			   int lastIndex,
			   char* buf);
void PrintIdeInfo (int drive, DWORD diskdata [256], LPSTR outputSerial);
BOOL DoIDENTIFY (HANDLE, PSENDCMDINPARAMS, PSENDCMDOUTPARAMS, BYTE, BYTE,
				 PDWORD);

bool ReadPhysicalDriveInNTWithAdminRights (int drive, LPSTR outputSerial)
{
   bool done = false;
   {
	  HANDLE hPhysicalDriveIOCTL = 0;

	 //  Try to get a handle to PhysicalDrive IOCTL, report failure
	 //  and exit if can't.
	  char driveName [256];

	  sprintf (driveName, "\\\\.\\PhysicalDrive%d", drive);

	 //  Windows NT, Windows 2000, must have admin rights
	  hPhysicalDriveIOCTL = CreateFileA (driveName,
							   GENERIC_READ | GENERIC_WRITE, 
							   FILE_SHARE_READ | FILE_SHARE_WRITE , NULL,
							   OPEN_EXISTING, 0, NULL);

	  if (hPhysicalDriveIOCTL != INVALID_HANDLE_VALUE)
	  {
		 GETVERSIONOUTPARAMS VersionParams;
		 DWORD               cbBytesReturned = 0;

		 // Get the version, etc of PhysicalDrive IOCTL
		 memset ((void*) &VersionParams, 0, sizeof(VersionParams));

		 if ( ! DeviceIoControl (hPhysicalDriveIOCTL, DFP_GET_VERSION,
				   NULL, 
				   0,
				   &VersionParams,
				   sizeof(VersionParams),
				   &cbBytesReturned, NULL) )
		 {
		 }

		 // If there is a IDE device at number "i" issue commands
		 // to the device
		 if (VersionParams.bIDEDeviceMap > 0)
		 {
			BYTE             bIDCmd = 0;   // IDE or ATAPI IDENTIFY cmd
			SENDCMDINPARAMS  scip;
			//SENDCMDOUTPARAMS OutCmd;

			// Now, get the ID sector for all IDE devices in the system.
			// If the device is ATAPI use the IDE_ATAPI_IDENTIFY command,
			// otherwise use the IDE_ATA_IDENTIFY command
			bIDCmd = (VersionParams.bIDEDeviceMap >> drive & 0x10) ? \
					  IDE_ATAPI_IDENTIFY : IDE_ATA_IDENTIFY;

			memset (&scip, 0, sizeof(scip));
			memset (IdOutCmd, 0, sizeof(IdOutCmd));

			if ( DoIDENTIFY (hPhysicalDriveIOCTL, 
					   &scip, 
					   (PSENDCMDOUTPARAMS)&IdOutCmd, 
					   (BYTE) bIDCmd,
					   (BYTE) drive,
					   &cbBytesReturned))
			{
			   DWORD diskdata [256];
			   int ijk = 0;
			   USHORT *pIdSector = (USHORT *)
							 ((PSENDCMDOUTPARAMS) IdOutCmd) -> bBuffer;

			   for (ijk = 0; ijk < 256; ijk++)
				  diskdata [ijk] = pIdSector [ijk];

			   PrintIdeInfo (drive, diskdata, outputSerial);
			   if (strlen(outputSerial) > 0)
				   done = true;
			}
		}

		 CloseHandle (hPhysicalDriveIOCTL);
	  }
   }

   return done;
}



//
// IDENTIFY data (from ATAPI driver source)
//

#pragma pack(1)

typedef struct _IDENTIFY_DATA {
	USHORT GeneralConfiguration;            // 00 00
	USHORT NumberOfCylinders;               // 02  1
	USHORT Reserved1;                       // 04  2
	USHORT NumberOfHeads;                   // 06  3
	USHORT UnformattedBytesPerTrack;        // 08  4
	USHORT UnformattedBytesPerSector;       // 0A  5
	USHORT SectorsPerTrack;                 // 0C  6
	USHORT VendorUnique1[3];                // 0E  7-9
	USHORT SerialNumber[10];                // 14  10-19
	USHORT BufferType;                      // 28  20
	USHORT BufferSectorSize;                // 2A  21
	USHORT NumberOfEccBytes;                // 2C  22
	USHORT FirmwareRevision[4];             // 2E  23-26
	USHORT ModelNumber[20];                 // 36  27-46
	UCHAR  MaximumBlockTransfer;            // 5E  47
	UCHAR  VendorUnique2;                   // 5F
	USHORT DoubleWordIo;                    // 60  48
	USHORT Capabilities;                    // 62  49
	USHORT Reserved2;                       // 64  50
	UCHAR  VendorUnique3;                   // 66  51
	UCHAR  PioCycleTimingMode;              // 67
	UCHAR  VendorUnique4;                   // 68  52
	UCHAR  DmaCycleTimingMode;              // 69
	USHORT TranslationFieldsValid:1;        // 6A  53
	USHORT Reserved3:15;
	USHORT NumberOfCurrentCylinders;        // 6C  54
	USHORT NumberOfCurrentHeads;            // 6E  55
	USHORT CurrentSectorsPerTrack;          // 70  56
	ULONG  CurrentSectorCapacity;           // 72  57-58
	USHORT CurrentMultiSectorSetting;       //     59
	ULONG  UserAddressableSectors;          //     60-61
	USHORT SingleWordDMASupport : 8;        //     62
	USHORT SingleWordDMAActive : 8;
	USHORT MultiWordDMASupport : 8;         //     63
	USHORT MultiWordDMAActive : 8;
	USHORT AdvancedPIOModes : 8;            //     64
	USHORT Reserved4 : 8;
	USHORT MinimumMWXferCycleTime;          //     65
	USHORT RecommendedMWXferCycleTime;      //     66
	USHORT MinimumPIOCycleTime;             //     67
	USHORT MinimumPIOCycleTimeIORDY;        //     68
	USHORT Reserved5[2];                    //     69-70
	USHORT ReleaseTimeOverlapped;           //     71
	USHORT ReleaseTimeServiceCommand;       //     72
	USHORT MajorRevision;                   //     73
	USHORT MinorRevision;                   //     74
	USHORT Reserved6[50];                   //     75-126
	USHORT SpecialFunctionsEnabled;         //     127
	USHORT Reserved7[128];                  //     128-255
} IDENTIFY_DATA, *PIDENTIFY_DATA;

#pragma pack()



bool ReadPhysicalDriveInNTUsingSmart (int drive, LPSTR outputSerial)
{
   bool done = false;

   {
	  HANDLE hPhysicalDriveIOCTL = 0;

		 //  Try to get a handle to PhysicalDrive IOCTL, report failure
		 //  and exit if can't.
	  char driveName [256];

	  sprintf (driveName, "\\\\.\\PhysicalDrive%d", drive);

		 //  Windows NT, Windows 2000, Windows Server 2003, Vista
	  hPhysicalDriveIOCTL = CreateFileA (driveName,
							   GENERIC_READ | GENERIC_WRITE, 
							   FILE_SHARE_DELETE | FILE_SHARE_READ | FILE_SHARE_WRITE, 
							   NULL, OPEN_EXISTING, 0, NULL);
	  // if (hPhysicalDriveIOCTL == INVALID_HANDLE_VALUE)
	  //    printf ("Unable to open physical drive %d, error code: 0x%lX\n",
	  //            drive, GetLastError ());

	  if (hPhysicalDriveIOCTL != INVALID_HANDLE_VALUE)
	  {
		 GETVERSIONINPARAMS GetVersionParams;
		 DWORD cbBytesReturned = 0;

			// Get the version, etc of PhysicalDrive IOCTL
		 memset ((void*) & GetVersionParams, 0, sizeof(GetVersionParams));

		 if (DeviceIoControl (hPhysicalDriveIOCTL, SMART_GET_VERSION,
				   NULL, 
				   0,
				   &GetVersionParams, sizeof (GETVERSIONINPARAMS),
				   &cbBytesReturned, NULL) )
		 {
			// Print the SMART version
			// PrintVersion (& GetVersionParams);
			// Allocate the command buffer
			ULONG CommandSize = sizeof(SENDCMDINPARAMS) + IDENTIFY_BUFFER_SIZE;
			PSENDCMDINPARAMS Command = (PSENDCMDINPARAMS) malloc (CommandSize);
			// Retrieve the IDENTIFY data
			// Prepare the command
#define ID_CMD          0xEC            // Returns ID sector for ATA
			Command -> irDriveRegs.bCommandReg = ID_CMD;
			DWORD BytesReturned = 0;
			if (DeviceIoControl (hPhysicalDriveIOCTL, 
									SMART_RCV_DRIVE_DATA, Command, sizeof(SENDCMDINPARAMS),
									Command, CommandSize,
									&BytesReturned, NULL) )
			{
				// Print the IDENTIFY data
				DWORD diskdata [256];
				USHORT *pIdSector = (USHORT *)
							 (PIDENTIFY_DATA) ((PSENDCMDOUTPARAMS) Command) -> bBuffer;

				for (int ijk = 0; ijk < 256; ijk++)
				   diskdata [ijk] = pIdSector [ijk];

				PrintIdeInfo (drive, diskdata, outputSerial);
				if (strlen(outputSerial) > 0)
					done = true;
			}
			// Done
			CloseHandle (hPhysicalDriveIOCTL);
			free (Command);
		 }
	  }
   }
   return done;
}

//  Required to ensure correct PhysicalDrive IOCTL structure setup
#pragma pack(4)


//
// IOCTL_STORAGE_QUERY_PROPERTY
//
// Input Buffer:
//      a STORAGE_PROPERTY_QUERY structure which describes what type of query
//      is being done, what property is being queried for, and any additional
//      parameters which a particular property query requires.
//
//  Output Buffer:
//      Contains a buffer to place the results of the query into.  Since all
//      property descriptors can be cast into a STORAGE_DESCRIPTOR_HEADER,
//      the IOCTL can be called once with a small buffer then again using
//      a buffer as large as the header reports is necessary.
//
#define IOCTL_STORAGE_QUERY_PROPERTY   CTL_CODE(IOCTL_STORAGE_BASE, 0x0500, METHOD_BUFFERED, FILE_ANY_ACCESS)

//  function to decode the serial numbers of IDE hard drives
//  using the IOCTL_STORAGE_QUERY_PROPERTY command 
char * flipAndCodeBytes (const char * str,
			 int pos,
			 int flip,
			 char * buf)
{
   int i;
   int j = 0;
   int k = 0;
   //Mr_TR---Status flip ---
   int trth=0;

   buf [0] = '\0';
   if (pos <= 0)
	  return buf;

   if ( ! j)
   {
	  char p = 0;

	  // First try to gather all characters representing hex digits only.
	  j = 1;
	  k = 0;
	  buf[k] = 0;
	  for (i = pos; j && str[i] != '\0'; ++i)
	  {

	 char c = tolower(str[i]);

	 if (isspace(c))
	 {
		c = '0';
	 }
	 ++p;
	 buf[k] <<= 4;

	 if (c >= '0' && c <= '9')
		buf[k] |= (unsigned char) (c - '0');
	 else if (c >= 'a' && c <= 'f')
		buf[k] |= (unsigned char) (c - 'a' + 10);
	 else
	 {
		j = 0;
		break;
	 }

	 if (p == 2)
	 {
		if (buf[k] != '\0' && ! isprint(buf[k]))
		{
		   j = 0;
		   break;
		}
		++k;
		p = 0;
		buf[k] = 0;
	 }

	  }
   }
   //Mr_TR--check buf null-------------------
   if(strcmp(buf,"")==0 ||strcmp(buf," ")==0)
   {
	   trth=1;
   }

   if ( ! j)
   {
	   
	  // There are non-digit characters, gather them as is.
	  j = 1;
	  k = 0;
	  for (i = pos; j && str[i] != '\0'; ++i)
	  {
		//VT--
		flip=0;
		 char c = str[i];

		 if ( ! isprint(c))
		 {
			j = 0;
			break;
		 }

		 buf[k++] = c;
	  }
   }

   if ( ! j)
   {
	  // The characters are not there or are not printable.
	  k = 0;
   }

   buf[k] = '\0';
   //VTR--------------
   char* bufp;
   bufp =  (char*) malloc( sizeof(char)*(int)strlen(buf) );
   strcpy(bufp,buf);
   // Flip adjacent characters
   if (flip || trth==1)
	  for (j = 0; j < k; j += 2)
	  {
		 char t = buf[j];
		 buf[j] = buf[j + 1];
		 buf[j + 1] = t;
	  }

   // Trim any beginning and end space
   i = j = -1;
   for (k = 0; buf[k] != '\0'; ++k)
   {
	  if (! isspace(buf[k]))
	  {
		 if (i < 0)
			i = k;
		 j = k;
	  }
   }

   if ((i >= 0) && (j >= 0))
   {
	  for (k = i; (k <= j) && (buf[k] != '\0'); ++k)
		 buf[k - i] = buf[k];
	  buf[k - i] = '\0';
   }

   if(stasDisk==1) strcpy(buf,bufp);

	return buf;
}


#define IOCTL_DISK_GET_DRIVE_GEOMETRY_EX CTL_CODE(IOCTL_DISK_BASE, 0x0028, METHOD_BUFFERED, FILE_ANY_ACCESS)

//typedef struct _DISK_GEOMETRY_EX {
//  DISK_GEOMETRY  Geometry;
//  LARGE_INTEGER  DiskSize;
//  UCHAR  Data[1];
//} DISK_GEOMETRY_EX, *PDISK_GEOMETRY_EX;

bool ReadPhysicalDriveInNTWithZeroRights (int drive, LPSTR outputSerial)
{
   bool done = false;
   {
	  HANDLE hPhysicalDriveIOCTL = 0;

	  //  Try to get a handle to PhysicalDrive IOCTL, report failure
	  //  and exit if can't.
	  char driveName [256];

	  sprintf (driveName, "\\\\.\\PhysicalDrive%d", drive);

	  //  Windows NT, Windows 2000, Windows XP - admin rights not required
	  hPhysicalDriveIOCTL = CreateFileA (driveName, 0,
							   FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
							   OPEN_EXISTING, 0, NULL);
	  if (hPhysicalDriveIOCTL != INVALID_HANDLE_VALUE)
	  {
		 STORAGE_PROPERTY_QUERY query;
		 DWORD cbBytesReturned = 0;
		 char buffer [10000];

		 memset ((void *) & query, 0, sizeof (query));
		 query.PropertyId = StorageDeviceProperty;
		 query.QueryType = PropertyStandardQuery;

		 memset (buffer, 0, sizeof (buffer));

		 if ( DeviceIoControl (hPhysicalDriveIOCTL, IOCTL_STORAGE_QUERY_PROPERTY,
				   & query,
				   sizeof (query),
				   & buffer,
				   sizeof (buffer),
				   & cbBytesReturned, NULL) )
		 {         
			 STORAGE_DEVICE_DESCRIPTOR * descrip = (STORAGE_DEVICE_DESCRIPTOR *) & buffer;
			 
			 char serialNumber [1000];
			 flipAndCodeBytes (buffer,
							   descrip -> SerialNumberOffset,
							   1, serialNumber );
			 strcpy(outputSerial, serialNumber);
			 if (strlen(outputSerial) > 0)
				 done = true;

		 }
		 CloseHandle (hPhysicalDriveIOCTL);
	  }
   }

   return done;
}


   // DoIDENTIFY
   // FUNCTION: Send an IDENTIFY command to the drive
   // bDriveNum = 0-3
   // bIDCmd = IDE_ATA_IDENTIFY or IDE_ATAPI_IDENTIFY
BOOL DoIDENTIFY (HANDLE hPhysicalDriveIOCTL, PSENDCMDINPARAMS pSCIP,
				 PSENDCMDOUTPARAMS pSCOP, BYTE bIDCmd, BYTE bDriveNum,
				 PDWORD lpcbBytesReturned)
{
	  // Set up data structures for IDENTIFY command.
   pSCIP -> cBufferSize = IDENTIFY_BUFFER_SIZE;
   pSCIP -> irDriveRegs.bFeaturesReg = 0;
   pSCIP -> irDriveRegs.bSectorCountReg = 1;
   //pSCIP -> irDriveRegs.bSectorNumberReg = 1;
   pSCIP -> irDriveRegs.bCylLowReg = 0;
   pSCIP -> irDriveRegs.bCylHighReg = 0;

	  // Compute the drive number.
   pSCIP -> irDriveRegs.bDriveHeadReg = 0xA0 | ((bDriveNum & 1) << 4);

	  // The command can either be IDE identify or ATAPI identify.
   pSCIP -> irDriveRegs.bCommandReg = bIDCmd;
   pSCIP -> bDriveNumber = bDriveNum;
   pSCIP -> cBufferSize = IDENTIFY_BUFFER_SIZE;

   return ( DeviceIoControl (hPhysicalDriveIOCTL, DFP_RECEIVE_DRIVE_DATA,
			   (LPVOID) pSCIP,
			   sizeof(SENDCMDINPARAMS) - 1,
			   (LPVOID) pSCOP,
			   sizeof(SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE - 1,
			   lpcbBytesReturned, NULL) );
}



#define  SENDIDLENGTH  sizeof (SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE

bool ReadIdeDriveAsScsiDriveInNT (int controller, LPSTR outputSerial)
{
   bool done = false;
   {
	  HANDLE hScsiDriveIOCTL = 0;
	  char   driveName [256];

	  //  Try to get a handle to PhysicalDrive IOCTL, report failure
	  //  and exit if can't.
	  sprintf (driveName, "\\\\.\\Scsi%d:", controller);

	  //  Windows NT, Windows 2000, any rights should do
	  hScsiDriveIOCTL = CreateFileA (driveName,
							   GENERIC_READ | GENERIC_WRITE, 
							   FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
							   OPEN_EXISTING, 0, NULL);

	  if (hScsiDriveIOCTL != INVALID_HANDLE_VALUE)
	  {
		 int drive = 0;

		 for (drive = 0; drive < 2; drive++)
		 {
			char buffer [sizeof (SRB_IO_CONTROL) + SENDIDLENGTH];
			SRB_IO_CONTROL *p = (SRB_IO_CONTROL *) buffer;
			SENDCMDINPARAMS *pin =
				   (SENDCMDINPARAMS *) (buffer + sizeof (SRB_IO_CONTROL));
			DWORD dummy;
   
			memset (buffer, 0, sizeof (buffer));
			p -> HeaderLength = sizeof (SRB_IO_CONTROL);
			p -> Timeout = 10000;
			p -> Length = SENDIDLENGTH;
			p -> ControlCode = IOCTL_SCSI_MINIPORT_IDENTIFY;
			strncpy ((char *) p -> Signature, "SCSIDISK", 8);
  
			pin -> irDriveRegs.bCommandReg = IDE_ATA_IDENTIFY;
			pin -> bDriveNumber = drive;

			if (DeviceIoControl (hScsiDriveIOCTL, IOCTL_SCSI_MINIPORT, 
								 buffer,
								 sizeof (SRB_IO_CONTROL) +
										 sizeof (SENDCMDINPARAMS) - 1,
								 buffer,
								 sizeof (SRB_IO_CONTROL) + SENDIDLENGTH,
								 &dummy, NULL))
			{
			   SENDCMDOUTPARAMS *pOut =
					(SENDCMDOUTPARAMS *) (buffer + sizeof (SRB_IO_CONTROL));
			   IDSECTOR *pId = (IDSECTOR *) (pOut -> bBuffer);
			   if (pId -> sModelNumber [0])
			   {
				  DWORD diskdata [256];
				  int ijk = 0;
				  USHORT *pIdSector = (USHORT *) pId;
		  
				  for (ijk = 0; ijk < 256; ijk++)
					 diskdata [ijk] = pIdSector [ijk];

				  PrintIdeInfo (controller * 2 + drive, diskdata, outputSerial);

				  if (strlen(outputSerial) > 0)
					  done = true;
			   }
			}
		 }
		 CloseHandle (hScsiDriveIOCTL);
	  }
   }

   return done;
}


void PrintIdeInfo (int drive, DWORD diskdata [256], LPSTR outputSerial)
{
   char serialNumber [MAX_LEN];
   char modelNumber [MAX_LEN];
   char revisionNumber [MAX_LEN];

	  //  copy the hard drive serial number to the buffer
   ConvertToString (diskdata, 10, 19, serialNumber);
   ConvertToString (diskdata, 27, 46, modelNumber);
   ConvertToString (diskdata, 23, 26, revisionNumber);

   strcpy(outputSerial, serialNumber);
}



char *ConvertToString (DWORD diskdata [256],
			   int firstIndex,
			   int lastIndex,
			   char* buf)
{
   int index = 0;
   int position = 0;

	  //  each integer has two characters stored in it backwards
   for (index = firstIndex; index <= lastIndex; index++)
   {
		 //  get high byte for 1st character
	  buf [position++] = (char) (diskdata [index] / 256);

		 //  get low byte for 2nd character
	  buf [position++] = (char) (diskdata [index] % 256);
   }

	  //  end the string 
   buf[position] = '\0';

	  //  cut off the trailing blanks
   for (index = position - 1; index > 0 && isspace(buf [index]); index--)
	  buf [index] = '\0';

   return buf;
}
