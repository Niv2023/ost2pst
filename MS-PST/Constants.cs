
namespace ost2pst
{
    public static class MS
    {
        //
        // constants values applicable to UNICODE
        public const UInt32 Magic = 0x4e444221;
        public const UInt16 MagicClient = 0x4d53;
        public const UInt16 minUnicodeVer = 0x17;
        public const UInt16 OSTversion = 0x24;      // OST file with 4k pages
        public const UInt16 VerClient = 19;
        public const UInt64 ibFileEofOffset = 184;
        public const UInt32 MaxBlockSize = 8192;
        public const UInt32 MaxDatablockSize = 8176; // block - block trailer
        public const UInt32 MaxXblockEntries = 1021; // (blocksize - 8)/8
        public const UInt32 MaxSLblockEntries = 340; // (blocksize - 8)/24
        public const int MaxHeapItemSize = 3580;   // stated in MS_PST 2.3.3.3 (PC) and 2.3.4.4.2 (TC)
        public const int MaxOSTBlockSize = 65536; 
        public const int BTPageEntryBytes = 488;
        public const int BTPAGEEntryBytesOST = 4056; // OST 4k page size
        public const int NBTcbEnt = 32;
        public const int BBTcbEnt = 24;
        public const int BTcbEnt = 24;
        public const int AMapBitAlloc = 64;
        public const int NodePageBlockSize = 512;
        public const int AMapPageEntryBytes = 496;
        public const uint AMapAllocationSize = 496 * 8 * 64;
        public const int DListPageEntryBytes = 476;
        public const int DListEntries = DListPageEntryBytes / 4; // Each DListEnt with 4 bytes
        public const UInt64 dListOffset = 0x4200;
        public const UInt64 firstAMapOffset = 0x4400;
        public const int MaxNrTreeLevels = 8;
    }
}