/*
 * Data Types specified in [MS-PST]-22115
 * - It implements only UNICODE PST/OST format
 * - only types required for replicating the NDB/BBT layers of a OST/PST
 * - in includes undocument OST specific types: https://blog.mythicsoft.com/ost-2013-file-format-the-missing-documentation/
 *      - needs to check whether the HID index change needs to be adjusted in the PST!!!!!!!
 * - PST's will be created with supports only permute crypto
 */
using System.Runtime.InteropServices;

namespace ost2pst
{
    public enum FileType { OST = 0, PST = 1 };
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public unsafe struct FileHeader
    {
        public UInt32 dwMagic;
        public UInt32 dwCRCPartial;
        public UInt16 wMagicClient;
        public UInt16 wVer;
        public UInt16 wVerClient;
        public Byte bPlatformCreate;
        public Byte bPlatformAccess;
        public UInt32 dwReserved1;
        public UInt32 dwReserved2;
        public UInt64 bidUnused;
        public UInt64 bidNextP;
        public UInt32 dwUnique;
        public fixed UInt32 rgnid[32];
        public UInt64 qwUnused;
        public Root root;
        public UInt32 dwAlign;
        public fixed Byte rgbFM[128];
        public fixed Byte rgbFP[128];
        public Byte bSentinel;
        public EbCryptMethod bCryptMethod;
        public UInt16 rgbReserved;
        public UInt64 bidNextB;
        public UInt32 dwCRCFull;
        public UInt32 rgbReserved2;
        public fixed Byte rgbReserved3[32];
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct Root
    {
        public UInt32 dwReserved;
        public UInt64 ibFileEof;
        public UInt64 ibAMapLast;
        public UInt64 cbAMapFree;
        public UInt64 cbPMapFree;
        public BREF BREFNBT;
        public BREF BREFBBT;
        public Byte fAMapValid;
        public Byte bReserved;
        public UInt16 wReserved;
    }
    public enum EbCryptMethod : byte
    {
        NDB_CRYPT_NONE = 0x00,
        NDB_CRYPT_PERMUTE = 0x01,
        NDB_CRYPT_CYCLIC = 0x02,
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct NID
    {

        public UInt32 dwValue; // References use the whole four bytes
        public EnidType nidType
        { // Low order five bits of stored value 
            get { return (EnidType)(dwValue & 0x0000001f); }
            set
            {
                dwValue &= 0xffffffe0;
                dwValue |= (UInt32)(value) & 0x0000001f;
            }
        }
        public UInt32 nidIndex
        {
            get { return dwValue >> 5; }
            set
            {
                dwValue &= 0x0000001f;
                dwValue |= (UInt32)(value << 5);
            }
        }
        public NID(EnidType type, UInt32 index)
        {
            dwValue |= (UInt32)(type) & 0x0000001f;
            dwValue |= (UInt32)(index << 5);
        }
        public NID(UInt32 nid)
        {
            this.dwValue = nid;
        }
        public NID ChangeType(EnidType type)
        {
            return new NID(type, nidIndex);
        }
        public static EnidType Type (UInt64 nid)
        {
            return (EnidType)(nid & 0x0000001f);
        }

    }
    public enum EnidType : byte
    {
        HID = 0x00,   // Heap node
        INTERNAL = 0x01,   // Internal node
        NORMAL_FOLDER = 0x02,  // Normal Folder object
        SEARCH_FOLDER = 0x03,  // Search Folder object
        NORMAL_MESSAGE = 0x04,
        ATTACHMENT = 0x05,
        SEARCH_UPDATE_QUEUE = 0x06,
        SEARCH_CRITERIA_OBJECT = 0x07,
        ASSOC_MESSAGE = 0x08,   // Folder associated information (FAI) Message object
        CONTENTS_TABLE_INDEX = 0x0a, // Internal, persisted view-related
        RECEIVE_FOLDER_TABLE = 0x0b, // Receive Folder object (Inbox)
        OUTGOING_QUEUE_TABLE = 0x0c, // Outbound queue (Outbox)
        HIERARCHY_TABLE = 0x0d,
        CONTENTS_TABLE = 0x0e,
        ASSOC_CONTENTS_TABLE = 0x0f,   // FAI contents table
        SEARCH_CONTENTS_TABLE = 0x10,   // Contents table (TC) of a search Folder object
        ATTACHMENT_TABLE = 0x11,
        RECIPIENT_TABLE = 0x12,
        SEARCH_TABLE_INDEX = 0x13,  // Internal, persisted view-related
        LTP = 0x1f,
    }
    public enum Eptype : byte
    {
        ptypeBBT = 0x80,   // Block BTH page
        ptypeNBT = 0x81,   // Node BTH page
        ptypeFMap = 0x82,  // Free Map page
        ptypePMap = 0x83,  // Allocation Page Map page
        ptypeAMap = 0x84,  // Allocation Map page
        ptypeFPMap = 0x85, // Free Page Map page
        ptypeDL = 0x86,    // Density List page
    }
    public enum EbType : byte
    {
        bTypeD = 0x00,    // just external data
        bTypeX = 0x01,   // XBLOCK or XXBLOCK
        bTypeS = 0x02,   // SLBLOCK or SIBLOCK
        bTypeTC = 0x7c,   // Table Context
        bTypeBTH = 0xb5,   // BTH-on-Heap
        bTypePC = 0xbc,  // Property Context
    }
    public enum EnidSpecial : uint
    {
        NID_MESSAGE_STORE = 0x0021,
        NID_NAME_TO_ID_MAP = 0x0061,
        NID_ROOT_FOLDER = 0x0122,
        NID_SEARCH_MANAGEMENT_QUEUE = 0x01E1,
        NID_SEARCH_ACTIVITY_LIST = 0x0201,
        NID_HIERARCHY_TABLE_TEMPLATE = 0x060D,
        NID_CONTENTS_TABLE_TEMPLATE = 0x060E,
        NID_ASSOC_CONTENTS_TABLE_TEMPLATE = 0x060F,
        NID_SEARCH_CONTENTS_TABLE_TEMPLATE = 0x0610,
        NID_RECIPIENT_TABLE = 0x0692,
        NID_ATTACHMENT_TABLE = 0x0671,
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct BREF
    {
        public UInt64 bid;
        public UInt64 ib;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe public struct BTPAGE
    {
        public fixed Byte rgentries[MS.BTPageEntryBytes];
        public Byte cEnt;
        public Byte cEntMax;
        public Byte cbEnt;
        public Byte cLevel;
        public UInt32 dwPadding;
        public PAGETRAILER pageTrailer;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct PAGETRAILER
    {
        public Eptype ptype;
        public Eptype ptypeRepeat;
        public UInt16 wSig;
        public UInt32 dwCRC;
        public UInt64 bid;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct BTENTRY
    {
        public UInt64 btkey;
        public BREF BREF;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct BBTENTRY
    {
        public BREF BREF;
        public UInt16 cb;
        public UInt16 cRef;
        public UInt32 dwPadding;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct NBTENTRY
    {
        public NID nid;                 // nid 4 bytes by only 2 in the NID type
        public UInt32 dwPad;            // thus dwPad -> 0
        public UInt64 bidData;
        public UInt64 bidSub;
        public UInt32 nidParent;
        public UInt32 dwPadding;
    }
    // LTP layer
    public enum EpropertyType : UInt16
    {
        PtypBinary = 0x0102,
        PtypBoolean = 0x000b,
        PtypFloating64 = 0x0005,
        PtypGuid = 0x0048,
        PtypInteger16 = 0x0002,
        PtypInteger32 = 0x0003,
        PtypInteger64 = 0x0014,
        PtypMultipleInteger32 = 0x1003,
        PtypObject = 0x000d,
        PtypString = 0x001f,
        PtypString8 = 0x001e,
        PtypMultipleString = 0x101F,
        PtypMultipleBinary = 0x1102,
        PtypTime = 0x0040,
    }
    public enum EpropertyId : UInt16
    {
        // Root folder
        PidTagRecordKey = 0x0FF9,
        PidTagIpmSubTreeEntryId = 0x35E0,
        PidTagIpmWastebasketEntryId = 0x35E3,
        PidTagFinderEntryId = 0x35E7,
        // Folder
        PidTagContentCount = 0x3602,
        PidTagContentUnreadCount = 0x3603,
        PidTagSubfolders = 0x360A,
        // Message List
        PidTagSubjectW = 0x0037,
        PidTagDisplayCcW = 0x0E03,
        PidTagDisplayToW = 0x0E04,
        PidTagMessageFlags = 0x0E07,
        PidTagMessageDeliveryTime = 0x0E06,
        PidTagReceivedByName = 0x0040,
        PidTagSentRepresentingNameW = 0x0042,
        PidTagSentRepresentingEmailAddress = 0x0065,
        PidTagSenderName = 0x0C1A,
        PidTagClientSubmitTime = 0x0039,
        PidTagCreationTime = 0x3007,
        PidTagLastModificationTime = 0x3008,
        // Message body
        PidTagNativeBody = 0x1016,
        PidTagBody = 0x1000,
        PidTagInternetCodepage = 0x3FDE,
        PidTagHtml = 0x1013,
        PidTagRtfCompressed = 0x1009,
        // Recipient
        PidTagRecipientType = 0x0c15,
        PidTagEmailAddress = 0x3003,
        // Attachment
        PidTagAttachFilenameW = 0x3704,
        PidTagAttachLongFilename = 0x3707,
        PidTagAttachmentSize = 0x0E20,
        PidTagAttachMethod = 0x3705,
        PidTagAttachMimeTag = 0x370e,
        PidTagAttachContentId = 0x3712,
        PidTagAttachFlags = 0x3714,
        PidTagAttachPayloadClass = 0x371a,
        PidTagAttachDataBinary = 0x3701,
        PidTagAttachmentHidden = 0x7ffe,
        //PidTagAttachDataObject = 0x3701,
        // Named properties
        PidTagNameidBucketCount = 0x0001,
        PidTagNameidStreamGuid = 0x0002,
        PidTagNameidStreamEntry = 0x0003,
        PidTagNameidStreamString = 0x0004,
        PidTagDisplayName = 0x3001,
        PidTagReplItemid = 0x0E30,
        PidTagReplChangenum = 0x0E33,
        PidTagReplVersionHistory = 0x0E34,
        PidTagReplFlags = 0x0E38,
        PidTagContainerClass = 0x3613,
        PidTagPstHiddenCount = 0x6635,
        PidTagPstHiddenUnread = 0x6636,
        PidTagLtpRowId = 0x67F2,
        PidTagLtpRowVer = 0x67F3,
        PidTagImportance = 0x0017,
        PidTagMessageClassW = 0x001A,
        PidTagSensitivity = 0x0036,
        PidTagMessageToMe = 0x0057,
        PidTagMessageCcMe = 0x0058,
        PidTagConversationTopicW = 0x0070,
        PidTagConversationIndex = 0x0071,
        PidTagMessageSize = 0x0E08,
        PidTagMessageStatus = 0x0E17,
        PidTagReplCopiedfromVersionhistory = 0x0E3C,
        PidTagReplCopiedfromItemid = 0x0E3D,
        PidTagItemTemporaryFlags = 0x1097,
        PidTagSecureSubmitFlags = 0x65C6,
        PidTagMessageClass = 0x001A,
        PidTagOfflineAddressBookName = 0x6800,
        PidTagSendOutlookRecallReport = 0x6803,
        PidTagOfflineAddressBookTruncatedProperties = 0x6805,
        PidTagViewDescriptorFlags = 0x7003,
        PidTagViewDescriptorLinkTo = 0x7004,
        PidTagViewDescriptorViewFolder = 0x7005,
        PidTagViewDescriptorName = 0x7006,
        PidTagViewDescriptorVersion = 0x7007,
        PidTagParentDisplayW = 0x0E05,
        PidTagExchangeRemoteHeader = 0x0E2A,
        PidTagLtpParentNid = 0x67F1,
        PidTagResponsibility = 0x0E0F,
        PidTagObjectType = 0x0FFE,
        PidTagEntryId = 0x0FFF,
        PidTagAddressType = 0x3002,
        PidTagSearchKey = 0x300B,
        PidTagDisplayType = 0x3900,
        PidTag7BitDisplayName = 0x39FF,
        PidTagSendRichInfo = 0x3A40,
        PidTagAttachSize = 0x0E20,
        PidTagRenderingPosition = 0x370B,
        PidTagValidFolderMask = 0x35DF,
        PidTagPstPassword = 0x67FF,
    }
    // The HID is a 32-bit value, with the following internal structure
    // non-4K: 5-bit Type; 11-bit Index; 16-bit BlockIndex
    // 4K:     5-bit Type; 14-bit Index; 13-bit BlockIndex
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct HID
    {
        private UInt16 wValue1;
        private UInt16 wValue2;

        private UInt16 hidIndex { get { return (UInt16)(wValue1 >> 5); } }
        private UInt16 hidBlockIndex { get { return wValue2; } }
        private UInt16 hidIndex4K { get { return (UInt16)((wValue1 >> 5) | ((wValue2 & 0x0007) << 11)); } }
        private UInt16 hidBlockIndex4K { get { return (UInt16)(wValue2 >> 3); } }


        public HID(UInt16 wValue1, UInt16 wValue2)
        {
            this.wValue1 = wValue1;
            this.wValue2 = wValue2;
        }
        public HID(HNID hnid)
        {
            wValue1 = (UInt16)(hnid.dwValue & 0xffff);
            wValue2 = (UInt16)(hnid.dwValue >> 16);
        }

        public HID(EnidType hidType, UInt16 hidIndex, UInt16 hidBlockIndex)
        {   // for pst unicode files
            wValue1 = (UInt16)((UInt16)hidType | (UInt16)(hidIndex << 5));
            wValue2 = hidBlockIndex;
        }
        public EnidType hidType { get { return (EnidType)(wValue1 & 0x001f); } }

        public UInt16 GetIndex(bool isUnicode4K)
        {
            return isUnicode4K ? hidIndex4K : hidIndex;
        }
        public UInt32 hidValue { get { return (UInt32)(wValue1 | wValue2 << 16); } }

        public UInt16 GetBlockIndex(bool isUnicode4K)
        {
            return isUnicode4K ? hidBlockIndex4K : hidBlockIndex;
        }
    }

    // A variation on HID used where the value can be either a HID or a NID
    // It is a HID iff hidType is EnidType.HID and the wValues are not both zero
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct HNID
    {
        private UInt16 wValue1;
        private UInt16 wValue2;

        private UInt16 hidIndex { get { return (UInt16)(wValue1 >> 5); } }
        private UInt16 hidBlockIndex { get { return wValue2; } }
        private UInt16 hidIndex4K { get { return (UInt16)((wValue1 >> 5) | ((wValue2 & 0x0007) << 11)); } }
        private UInt16 hidBlockIndex4K { get { return (UInt16)(wValue2 >> 3); } }

        public bool HasValue { get { return (wValue1 != 0) || (wValue2 != 0); } }
        public bool IsHID { get { return HasValue && hidType == EnidType.HID; } }
        public bool IsNID => !IsHID;
        public EnidType hidType { get { return (EnidType)(wValue1 & 0x001f); } }
        public EnidType nidType { get { return (EnidType)(wValue1 & 0x001f); } }  // Low order five bits of stored value
        public UInt32 dwValue { get { return (UInt32)(wValue2 << 16) | wValue1; } }  // References use the whole four bytes
        public HID HID { get { return new HID(wValue1, wValue2); } }
        public NID NID { get { return new NID(dwValue); } }

        public HNID(HID hid)
        {
            wValue1 = (UInt16)(hid.hidValue & 0xffff);
            wValue2 = (UInt16)(hid.hidValue >> 16);
        }
        public HNID(NID nid)
        {
            wValue1 = (UInt16)(nid.dwValue & 0xffff);
            wValue2 = (UInt16)(nid.dwValue >> 16);
        }
        public HNID(UInt32 dwValue)
        {
            wValue1 = (UInt16)(dwValue & 0xffff);
            wValue2 = (UInt16)(dwValue >> 16);
        }
        public UInt16 GetIndex(bool isUnicode4K)
        {
            return isUnicode4K ? hidIndex4K : hidIndex;
        }

        public UInt16 GetBlockIndex(bool isUnicode4K)
        {
            return isUnicode4K ? hidBlockIndex4K : hidBlockIndex;
        }
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct HNHDR
    {
        public UInt16 ibHnpm;       // The byte offset to the HN page Map record
        public Byte bSig;
        public EbType bClientSig;
        public HID hidUserRoot;  // HID that points to the User Root record
        public UInt32 rgbFillLevel;
    }
    public struct PCBTH
    {
        public EpropertyId wPropId;
        public EpropertyType wPropType;
        public HNID dwValueHnid;
    }
    //
    // Table Context structures
    //

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct TCINFO
    {
        public EbType btype;    // Must be bTypeTC
        public Byte cCols;
        public UInt16 rgibTCI_4b;
        public UInt16 rgibTCI_2b;
        public UInt16 rgibTCI_1b;
        public UInt16 rgibTCI_bm;
        public HID hidRowIndex;
        public HNID hnidRows;
        public HID hidIndex;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public class TCOLDESC
    {
        public UInt32 tag;
        public EpropertyId wPropId { get { return (EpropertyId)(tag >> 16); } }
        public EpropertyType wPropType { get { return (EpropertyType)(tag & 0x0000ffff); } }
        public UInt16 ibData;
        public Byte cbData;
        public Byte iBit;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct TCROWID
    {
        public UInt32 dwRowID;
        public UInt32 dwRowIndex;
    }
    public class RowData
    {
        public List<Property> Props;
        public Byte[] rgbCEB;
        public Byte[] rgbData;
        public byte[] Array
        {
            get
            {
                byte[] data = new byte[rgbData.Length + rgbCEB.Length];
                rgbData.CopyTo(data, 0);
                rgbCEB.CopyTo(data, rgbData.Length);
                return data;
            }
        }
        public RowData()
        {

        }
        public RowData(TCINFO tcInfo, UInt32 rowId)
        {
            //          _cells = new List<TableContextCell>();
            rgbData = new Byte[tcInfo.rgibTCI_1b];
            rgbCEB = new Byte[tcInfo.rgibTCI_bm - tcInfo.rgibTCI_1b];
            FM.TypeToArray<UInt32>(ref rgbData, rowId);
            FM.TypeToArray<UInt32>(ref rgbData, FM.NextUniqueID(), 4);
        }
        public RowData(TCINFO tcInfo, UInt32 rowId, List<Property> props)
        {
            rgbData = new Byte[tcInfo.rgibTCI_1b];
            rgbCEB = new Byte[tcInfo.rgibTCI_bm - tcInfo.rgibTCI_1b];
            FM.TypeToArray<UInt32>(ref rgbData, rowId);
            FM.TypeToArray<UInt32>(ref rgbData, FM.NextUniqueID(), 4);
            PST.SetBitInArray(ref rgbCEB, 0);
            PST.SetBitInArray(ref rgbCEB, 1);
            Props = props;
        }
    }

    public class TableContext : ICloneable
    {
        private ushort lastIbdata;
        public TCINFO tcINFO;
        public List<TCOLDESC> tcCols;
        public List<TCROWID> tcRowIndexes;
        public List<RowData> tcRowMatrix;
        const UInt32 tagLtpRowId = (UInt32)EpropertyId.PidTagLtpRowId << 16 | (UInt32)EpropertyType.PtypInteger32;
        const UInt32 tagLtpRowVer = (UInt32)EpropertyId.PidTagLtpRowVer << 16 | (UInt32)EpropertyType.PtypInteger32;
        public TableContext()
        {
            lastIbdata = 0;
            tcINFO = new TCINFO()
            {
                btype = EbType.bTypeTC,
                hidRowIndex = new HID(EnidType.HID, 1, 0),            // outlook differs from the MS-PST
                hnidRows = new HNID(0),
                hidIndex = new HID(),
            };
            tcCols = new List<TCOLDESC>();
            tcRowIndexes = new List<TCROWID>();
            tcRowMatrix = new List<RowData>();
        }
        public object Clone()
        {
            return this.MemberwiseClone();
        }
        public void resetRows()
        {
            tcRowIndexes = new List<TCROWID>();
            tcRowMatrix = new List<RowData>();
        }
        public int RowMatrixSize
        {
            get
            {
                return tcRowMatrix.Count * tcINFO.rgibTCI_bm;
            }
        }
        public void AddColumn(EpropertyId propId, EpropertyType propType)
        {
            UInt32 pTag = (UInt32)propId << 16 | (UInt32)propType;
            // if (tcCols.FindIndex(x => x.tag == pTag) >= 0) { throw new Exception($"Column {propId} already defined);"); }
            if (FindColumn(propId, propType) >= 0) { throw new Exception($"Column {propId} already defined);"); }
            byte cb = GetPropertyTypeSize(propType);
            if (cb == 0) { cb = 4; } // variable size. cell will hold the HNID
                                     // overwriting to ensure RowId and RowVer are the first 2 cols -> MS_PST 2.3.4.4.1
            TCOLDESC tCOLDESC = new TCOLDESC()
            {
                tag = pTag,
                ibData = 0,
                cbData = cb,
                iBit = 0
            };
            tcCols.Add(tCOLDESC);
            tcINFO.cCols++;
            sortColumns();
            AddPropertyToRowData(propId, propType);
        }
        public void RemoveColumn(EpropertyId propId, EpropertyType propType)
        {
            UInt32 pTag = (UInt32)propId << 16 | (UInt32)propType;
            int colIx = tcCols.FindIndex(x => x.tag == pTag);
            if (colIx >= 0)
            {
                tcCols.RemoveAt(colIx);
                tcINFO.cCols--;
                sortColumns();
            };
            RemovePropertyToRowData(propId, propType);
        }
        private void RemovePropertyToRowData(EpropertyId propId, EpropertyType propType)
        {   //******************* need to check order/IX rowix  ---
            Property prop = new Property(propId, propType);
            for (int i = 0; i < tcRowMatrix.Count; i++)
            {
                List<Property> props = tcRowMatrix[i].Props;
                /*** code to be improved... duplication....... */
                int propIx = props.FindIndex(f => f.id == propId);
                if (propIx >= 0)
                {
                   props.RemoveAt(propIx);
                }
                UInt32 rowId = tcRowIndexes[i].dwRowID;
                tcRowMatrix[i] = new RowData(tcINFO, rowId, props);
            }
        }

        private void AddPropertyToRowData(EpropertyId propId, EpropertyType propType)
        {   //******************* need to check order/IX rowix
            Property prop = new Property(propId, propType);
            for (int i = 0; i < tcRowMatrix.Count; i++)
            {
                List<Property> props = tcRowMatrix[i].Props;
                //props.Add(prop);
                UInt32 rowId = tcRowIndexes[i].dwRowID;
                tcRowMatrix[i] = new RowData(tcINFO, rowId, props);
            }
        }
        public int FindColumn(EpropertyId propId, EpropertyType propType)
        {
            UInt32 pTag = (UInt32)propId << 16 | (UInt32)propType;
            int result = tcCols.FindIndex(x => x.tag == pTag);
            return result;
        }
        public void AddRow(TCROWID rId, List<Property> props)
        {
            // tcRowIndex already ordered 
            RowData rowData = new RowData(tcINFO, rId.dwRowID, props);
            tcRowMatrix.Add(rowData);
        }
        public void AddRow(NID nid, List<Property> props)
        {
            TCROWID rowId = new TCROWID()
            {
                dwRowID = nid.dwValue,
                dwRowIndex = (UInt32)tcRowIndexes.Count
            };
            tcRowIndexes.Add(rowId);
            tcRowIndexes = tcRowIndexes.OrderBy(x => x.dwRowID).ToList();
            RowData rowData = new RowData(tcINFO, nid.dwValue, props);
            tcRowMatrix.Add(rowData);
        }
        public unsafe List<byte[]> RowMatrixArray()
        {
            List<byte[]> rmBlocks = new List<byte[]>();
            int nrRows = tcRowMatrix.Count;
            int rowSize = tcINFO.rgibTCI_bm;
            int rowsPerBlock = (int)MS.MaxDatablockSize / rowSize;
            int nrBlocks = (int)Math.Ceiling((decimal)tcRowMatrix.Count / rowsPerBlock);
            int blockIx = 0;
            byte[] rowMatrix = new byte[0];
            for (int i = 0; i < nrRows; i++)
            {
                if (rowMatrix.Length + rowSize > MS.MaxDatablockSize)
                {  // closes the current block
                    Array.Resize(ref rowMatrix, (int)MS.MaxDatablockSize);
                    rmBlocks.Add(rowMatrix);
                    rowMatrix = new byte[0];
                    blockIx++;
                }
                Array.Resize(ref rowMatrix, rowMatrix.Length + rowSize);
                Array.Copy(tcRowMatrix[i].Array, 0, rowMatrix, rowSize * (i - blockIx * rowsPerBlock), rowSize);
            }
            if (rowMatrix.Length > 0) { rmBlocks.Add(rowMatrix); }
            return rmBlocks;
        }
        public unsafe byte[] Header()
        { // connot add new cols afterwards... 
            int tcInfoHeader = sizeof(TCINFO);
            int tcColDesc = sizeof(TCOLDESC);
            byte[] headerArray = new byte[tcInfoHeader + tcColDesc * tcCols.Count];
            FM.TypeToArray<TCINFO>(ref headerArray, tcINFO);
            for (int i = 0; i < tcCols.Count; i++)
            {
                if (tcCols[i].tag == tagLtpRowId | tcCols[i].tag == tagLtpRowVer) { tcCols[i].cbData = 4; }  // interger32
                FM.TypeToArray<TCOLDESC>(ref headerArray, tcCols[i], tcInfoHeader + tcColDesc * i);
            }
            return headerArray;
        }
        private byte GetPropertyTypeSize(EpropertyType propType)
        {   // variable (unknown) types returns 0
            byte cb = 0;
            switch (propType)
            {
                case EpropertyType.PtypInteger16:
                    cb = sizeof(UInt16);
                    break;
                case EpropertyType.PtypInteger32:
                    cb = sizeof(UInt32);
                    break;
                case EpropertyType.PtypInteger64:
                    cb = sizeof(UInt64);
                    break;
                case EpropertyType.PtypFloating64:
                    cb = sizeof(Double);
                    break;
                case EpropertyType.PtypBoolean:
                    cb = sizeof(bool);
                    break;
                case EpropertyType.PtypTime:
                    cb = sizeof(UInt64);
                    break;
                    // variable size types:
                    // - EpropertyType.PtypMultipleInteger32:
                    // - EpropertyType.PtypBinary:
                    // - EpropertyType.PtypString: // Unicode string
                    // - EpropertyType.PtypString8:  // Multipoint string in variable encoding
                    // - EpropertyType.PtypMultipleString: // Unicode strings
                    // - EpropertyType.PtypMultipleBinary:
                    // - EpropertyType.PtypGuid:
                    // - EpropertyType.PtypObject:
            }
            return cb;
        }
        private void sortColumns()
        {
            // sort desceding on cbtata
            UInt16 ibData = 8;
            byte iBit = 2;
            tcCols = tcCols.OrderByDescending(h => h.cbData).ToList();
            for (int i = 0; i < tcCols.Count; i++)
            {
                tcCols[i].ibData = ibData;
                tcCols[i].iBit = iBit;
                if (tcCols[i].tag == tagLtpRowId)
                {
                    tcCols[i].ibData = 0;
                    tcCols[i].iBit = 0;
                }
                else
                {
                    if (tcCols[i].tag == tagLtpRowVer)
                    {
                        tcCols[i].ibData = 4;
                        tcCols[i].iBit = 1;
                    }
                    else
                    {
                        ibData += tcCols[i].cbData;
                        iBit++;
                    }
                }
            }
            tcINFO.rgibTCI_bm = ibData;
            UInt16 cbCeb = (UInt16)(tcCols.Count / 8);
            if (tcCols.Count % 8 > 0) cbCeb++;
            tcINFO.rgibTCI_bm += cbCeb;
            tcINFO.rgibTCI_4b = tcINFO.rgibTCI_2b = tcINFO.rgibTCI_1b = 0;
            UInt16 offset = 0;
            for (int i = 0; i < tcCols.Count; i++)
            {
                //   offset = (ushort)(tcCols[i].ibData + tcCols[i].cbData);
                offset += tcCols[i].cbData;
                if (tcCols[i].cbData >= 4) tcINFO.rgibTCI_4b = tcINFO.rgibTCI_4b = tcINFO.rgibTCI_2b = tcINFO.rgibTCI_1b = offset;
                else if (tcCols[i].cbData == 2) tcINFO.rgibTCI_2b = tcINFO.rgibTCI_1b = offset;
                else if (tcCols[i].cbData == 1) tcINFO.rgibTCI_1b = offset;
            }
            tcCols = tcCols.OrderBy(h => h.tag).ToList();
        }

    }
    //
    // Page structures
    //
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct HNPAGEHDR
    {
        public UInt16 ibHnpm;       // The byte offset to the HN page Map record
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct HNBITMAPHDR
    {
        public UInt16 ibHnpm;       // The byte offset to the HN page Map record
        public fixed Byte rgbFillLevel[64];
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct HNPAGEMAP
    {
        public UInt16 cAlloc;
        public UInt16 cFree;
        // Marshal the following array manually
        //public UInt16[] rgibAlloc;
    }
    //
    // BTH-on-Heap structures
    //
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct BTHHEADER
    {
        public EbType btype;
        public Byte cbKey;
        public Byte cbEnt;
        public Byte bIdxLevels;
        public HID hidRoot;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct BTHENTRY // ==> used as default intermediate on the output pst...
    {
        public UInt32 key;
        public HID hidNextLevel;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct IntermediateBTH2
    {
        public fixed Byte key[2];
        public HID hidNextLevel;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct IntermediateBTH4
    {
        public fixed Byte key[4];
        public HID hidNextLevel;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct PtypObjectValue
    {
        public UInt32 Nid;
        public UInt32 ulSize;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct BLOCKTRAILER
    {
        public UInt16 cb;
        public UInt16 wSig;
        public UInt32 dwCRC;
        public UInt64 bid;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct XBLOCK
    {
        public EbType btype;
        public Byte cLevel;
        public UInt16 cEnt;
        public UInt32 lcbTotal;
        // Marshal the following array manually
        //public UInt64[] rgbid;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct SLENTRY
    {
        public UInt64 nid;
        public UInt64 bidData;
        public UInt64 bidSub;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct SLBLOCK
    {
        public EbType btype;
        public Byte cLevel;
        public UInt16 cEnt;
        public UInt32 dwPadding;
        // Marshal the following array manually
        //public SLENTRYUnicode[] rgentries;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct SIENTRY
    {
        public UInt64 nid;
        public UInt64 bid;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public unsafe struct AMapPage
    {
        public fixed byte rgbAMapBits[MS.AMapPageEntryBytes];
        public PAGETRAILER pageTrailer;
    }

    /*
     * PMapPage, FMapPage and FPMapPages are deprecated (replaced by DList)
     * but will be added in the PST for backwards compatibility (file structure)
     */
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct PMapPage
    {
        public fixed byte rgbPMapBits[MS.AMapPageEntryBytes];
        public PAGETRAILER pageTrailer;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct FMapPage
    {
        public fixed byte rgbFMapBits[MS.AMapPageEntryBytes];
        public PAGETRAILER pageTrailer;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct FPMapPage
    {
        public fixed byte rgbFPMapBits[MS.AMapPageEntryBytes];
        public PAGETRAILER pageTrailer;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct DLISTPAGEENT
    {
        public UInt32 dwValue; // References use the whole four bytes
        public uint dwPageNum
        { // 20 least significant bits
            get { return (uint)(dwValue & 0x000fffff); }
            set
            {
                dwValue &= 0xfff00000;
                dwValue |= (UInt32)(value & 0x000fffff);
            }
        }
        public uint dwFreeSlots
        { // 12 most significant bits
            get { return (uint)(dwValue >> 20); }
            set
            {
                dwValue &= 0x000fffff;
                dwValue |= (UInt32)(value << 20);
            }
        }
        public UInt32 Value
        {
            get { return dwValue; }
        }
        public DLISTPAGEENT(UInt32 dListEntry)
        {
            dwValue = dListEntry;
        }
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public unsafe struct DLISTPAGEENTRIES
    {
        public fixed UInt32 dwValue[MS.DListEntries]; // 
        public DLISTPAGEENT this[int i]
        {
            get { return new DLISTPAGEENT(dwValue[i]); }
            set { dwValue[i] = value.Value; }
        }
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public unsafe struct DListPage
    {
        public Byte bFlags;
        public Byte cEntDList;
        public UInt16 wPadding;
        public UInt32 ulCurrentPage;
        public DLISTPAGEENTRIES rgDListPageEntries;
        public fixed byte rgPadding[12];
        public PAGETRAILER pageTrailer;
    }
    // used in a AMap list for DList
    public class AMap : IComparable<AMap>
    {
        public int index;  // zero based
        public UInt64 offset;
        public uint nrFreeSlots; // 64 bytes per slot 
        public uint freeSpace { get { return nrFreeSlots * 64; } }
        public int CompareTo(AMap? compareAMap)
        {
            // A null value means that this object is greater.
            if (compareAMap == null)
                return 1;
            else
                return this.nrFreeSlots.CompareTo(compareAMap.nrFreeSlots);
        }
    }

    /*
    * OST data types
    */

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe struct OST_BTPAGE
    {
        public fixed Byte rgentries[MS.BTPAGEEntryBytesOST];
        public UInt16 cEnt;
        public UInt16 cEntMax;
        public Byte cbEnt;
        public Byte cLevel;
        public UInt16 dwPadding1;
        public UInt32 dwPadding2;
        public UInt32 dwPadding3;
        public OST_PAGETRAILER pageTrailer;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    struct OST_PAGETRAILER
    {
        public Eptype ptype;
        public Eptype ptypeRepeat;
        public UInt16 wSig;
        public UInt32 dwCRC;
        public UInt64 bid;
        public UInt64 unknown;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct OST_BBTENTRY
    {
        public BREF BREF;
        public UInt16 cbStored;
        public UInt16 cbInflated;
        public UInt16 cRef;
        public UInt16 wPadding;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct OST_BLOCKTRAILER
    {
        public UInt16 cb;
        public UInt16 wSig;
        public UInt32 dwCRC;
        public UInt64 bid;
        public UInt16 unknown1;
        public UInt16 cbInflated;
        public UInt32 unknown2;
    }
}