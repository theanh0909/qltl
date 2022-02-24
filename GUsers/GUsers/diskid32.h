bool ReadIdeDriveAsScsiDriveInNT (int controller, LPSTR outputSerial);
bool ReadPhysicalDriveInNTUsingSmart (int drive, LPSTR outputSerial);
bool ReadPhysicalDriveInNTWithAdminRights (int drive, LPSTR outputSerial);
bool ReadPhysicalDriveInNTWithZeroRights (int drive, LPSTR outputSerial);
