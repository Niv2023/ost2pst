namespace ost2pst
{
    public static class NDB
    {
        static EnidType[] pcNidTypes = new[] { EnidType.NORMAL_FOLDER, EnidType.SEARCH_FOLDER, EnidType.NORMAL_MESSAGE, EnidType.ATTACHMENT, EnidType.ASSOC_MESSAGE };
        static UInt32[] pcNIDs = new UInt32[] { 0x21, 0x61, 0x122 };
        static EnidType[] tcNidTypes = new[] { EnidType.HIERARCHY_TABLE, EnidType.CONTENTS_TABLE, EnidType.ASSOC_CONTENTS_TABLE, EnidType.SEARCH_CONTENTS_TABLE,
                                        EnidType.ATTACHMENT_TABLE, EnidType.RECIPIENT_TABLE, (EnidType)22};
        static UInt32[] tcNIDs = new UInt32[] { 0xA1, 0xC1};

        static public bool IsPC(NID nid)
        {
            if (nid.nidType == EnidType.INTERNAL)
            {
                return pcNIDs.Contains(nid.dwValue);
            }
            else
            {
                return pcNidTypes.Contains(nid.nidType);
            }
        }
        static public bool IsTC(NID nid)
        {
            if (nid.nidType == EnidType.INTERNAL)
            {
                return tcNIDs.Contains(nid.dwValue);
            }
            else
            {
                return tcNidTypes.Contains(nid.nidType);
            }
        }
    }
}
