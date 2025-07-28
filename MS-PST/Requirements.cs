/* 
 *  Functions to built PST objects as specified on 2.7 Minimum PST Requirements
 *  - new clean template objects will be created for the new PST file
 *  - however some may need to be dupIt implements only UNICODE PST/OST format
 * 
 */

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ost2pst
{
    public static class RQS
    {

        // NIDs
        public static NID NID_MESSAGE_STORE = new NID(0x0021);
        public static NID NID_NAME_TO_ID_MAP = new NID(0x0061);
        public static NID NID_ROOT_FOLDER = new NID(0x0122);
        public static NID NID_SEARCH_MANAGEMENT_QUEUE = new NID(0x01E1);
        public static NID NID_SEARCH_ACTIVITY_LIST = new NID(0x0201);
        public static NID NID_HIERARCHY_TABLE_TEMPLATE = new NID(0x060D);
        public static NID NID_CONTENTS_TABLE_TEMPLATE = new NID(0x060E);
        public static NID NID_ASSOC_CONTENTS_TABLE_TEMPLATE = new NID(0x060F);
        public static NID NID_SEARCH_CONTENTS_TABLE_TEMPLATE = new NID(0x0610);
        public static NID NID_OUTGOING_QUEUE_TABLE = new NID(0x064C);
        public static NID NID_RECIPIENT_TABLE = new NID(0x0692);
        public static NID NID_RECIEVE_FOLDER_TABLE = new NID(0x062B);
        public static NID NID_ATTACHMENT_TABLE = new NID(0x0671);
        public static NID NID_SPAM_SEARCH_FOLDER = new NID(0x2223);
        public static NID NID_TOP_PERSONAL_FOLDER = new NID(0x8022);
        public static NID NID_SEARCH_FOLDER = new NID(0x8042);
        public static NID NID_DELETED_ITEMS = new NID(0x8062);

        // TC templates
        public static TableContext HIERARCHY_TABLE_TEMPLATE;
        public static UInt64 HIERARCHY_TABLE_TEMPLATE_BID;
        public static TableContext CONTENTS_TABLE_TEMPLATE;
        public static UInt64 CONTENTS_TABLE_TEMPLATE_BID;
        public static TableContext ASSOC_CONTENTS_TABLE_TEMPLATE;
        public static UInt64 ASSOC_CONTENTS_TABLE_TEMPLATE_BID;
        public static TableContext SEARCH_CONTENTS_TABLE_TEMPLATE;
        public static UInt64 SEARCH_CONTENTS_TABLE_TEMPLATE_BID;
        public static TableContext RECIPIENT_TABLE_TEMPLATE;
        public static UInt64 RECIPIENT_TABLE_TEMPLATE_BID;
        public static TableContext ATTACHMENT_TABLE_TEMPLATE;
        public static UInt64 ATTACHMENT_TABLE_TEMPLATE_BID;


        public static void BuildPSTobjects(Folder folderToExport, string filename)
        { // 2.7.3	Minimum Object Requirements
          // - Creates a new Message Store
          // - Validates OST templates and creates new if required
            PSTMessageStore(filename);
            ValidateNameToIDmap();
            PSTTemplateObjects();
            CreatePSTfolderStructure(folderToExport);
        }
        private static void PSTMessageStore(string dispName)
        {   // Message Store properties  -> MS_PST 2.4.3
            byte[] rgbFlags = new byte[4] { 0, 0, 0, 0 };
            //byte[] pstUID = new byte[16] { 0xf0, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa, 0xfb, 0xfc, 0xfd, 0xfe, 0xff };
            byte[] pstUID = new byte[16] { 0x45, 0x7C, 0x5C, 0x24, 0x7F, 0x65, 0xF5, 0x45, 0x87, 0xF8, 0xD6, 0xA3, 0x0D, 0x30, 0xC2, 0x69 };
            byte[] entryPrefix = rgbFlags.Concat(pstUID).ToArray();
            byte[] subTreeEntry = entryPrefix.Concat(NidToArray(NID_TOP_PERSONAL_FOLDER)).ToArray();
            byte[] wasteBasketEntry = entryPrefix.Concat(NidToArray(NID_DELETED_ITEMS)).ToArray();
            byte[] finderEntry = entryPrefix.Concat(NidToArray(NID_SEARCH_FOLDER)).ToArray();
            // above properties mandatory according to 2.4.3
            // properties below show in a empty pst created in outlook:
            Int32 folderMask = 0x0089;
            Int32 flags = 0;
            Int32 password = 0;
            //  byte[] verHis = new byte[24] { 0x01, 0x00, 0x00, 0x00, 0xF3, 0x88, 0x89, 0x73, 0x0A, 0x45, 0xD1, 0x46,
            //                                 0x97, 0xC3, 0x17, 0xF5, 0x07, 0x27, 0x32, 0x38, 0x01, 0x00, 0x00, 0x00 };
//            24 bytes[01.00.00.00.36.FE.ED.74.58.4F.CE.4A.95.6F.75.23.08.BE.AA.C1.01.00.00.00]

            List<Property> props = new List<Property>
            {
                SetPropertyValue(EpropertyId.PidTagRecordKey, EpropertyType.PtypBinary, pstUID),
                SetPropertyValue(EpropertyId.PidTagDisplayName, EpropertyType.PtypString, dispName),
                SetPropertyValue(EpropertyId.PidTagIpmSubTreeEntryId, EpropertyType.PtypBinary, subTreeEntry),
                SetPropertyValue(EpropertyId.PidTagIpmWastebasketEntryId, EpropertyType.PtypBinary, wasteBasketEntry),
                SetPropertyValue(EpropertyId.PidTagFinderEntryId, EpropertyType.PtypBinary, finderEntry),
                //
                // from empty pst
                // SetPropertyValue(EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary, verHis),
                SetPropertyValue(EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary, LTP.ReplVersionHistory),
                SetPropertyValue(EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32, flags),
                SetPropertyValue(EpropertyId.PidTagValidFolderMask, EpropertyType.PtypInteger32, folderMask),
                SetPropertyValue(EpropertyId.PidTagPstPassword, EpropertyType.PtypInteger32, password)   
            };
            props = props.OrderBy(h => h.PCBTH.wPropId).ToList();
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            UInt64 newBid = FM.outFile.AddNDBdata(NID_MESSAGE_STORE, pcHNblocks, 0, 0);
        }
        private static void ValidateNameToIDmap()
        {   // will use the the OST nameToIDmap
            int nIx = FM.srcFile.NBTs.FindIndex(n => n.nid.dwValue == (UInt32)EnidSpecial.NID_NAME_TO_ID_MAP);
            if (nIx < 0) throw new Exception($"Missing mandatory Name-to-ID NID (0x61)");
        }
        private static void PSTTemplateObjects()
        {
            HIERARCHY_TABLE_TEMPLATE_BID = AddNIDtc(NID_HIERARCHY_TABLE_TEMPLATE, CREATE_HIERARCHY_TABLE_TEMPLATE());
            CONTENTS_TABLE_TEMPLATE_BID = AddNIDtc(NID_CONTENTS_TABLE_TEMPLATE, CREATE_CONTENTS_TABLE_TEMPLATE());
            ASSOC_CONTENTS_TABLE_TEMPLATE_BID = AddNIDtc(NID_ASSOC_CONTENTS_TABLE_TEMPLATE, CREATE_ASSOC_CONTENTS_TABLE_TEMPLATE());
            SEARCH_CONTENTS_TABLE_TEMPLATE_BID = AddNIDtc(NID_SEARCH_CONTENTS_TABLE_TEMPLATE, CREATE_SEARCH_CONTENTS_TABLE_TEMPLATE());
            RECIPIENT_TABLE_TEMPLATE_BID = AddNIDtc(NID_RECIPIENT_TABLE, CREATE_RECIPIENT_TABLE_TEMPLATE());
            ATTACHMENT_TABLE_TEMPLATE_BID = AddNIDtc(NID_ATTACHMENT_TABLE, CREATE_ATTACHMENT_TABLE_TEMPLATE());
            AddSearcActivityObject();
            AddReceiveFolderTable();
            AddOutgoingQueueTable();
            AddUndocObjects();
        }
        private static void CreatePSTfolderStructure(Folder exportFolder)
        {   // MS_PST 2.7.2.4	Folders
            //
            // 2.7.3.4.1	Root Folder
            Folder rootFolder = FolderPC(NID_ROOT_FOLDER, "", NID_ROOT_FOLDER, true);
            //
            // 2.7.3.4.2	Top of Personal Folders (IPM SuBTree)
            //List<Property> topFolderReplProps = replicationItems(0x0160, 0x1fA4, 0x0140);
            //Folder topPersonal = FolderPC(NID_TOP_PERSONAL_FOLDER, "Top of Personal Folders", NID_ROOT_FOLDER, true, topFolderReplProps);
            Folder topPersonal = FolderPC(NID_TOP_PERSONAL_FOLDER, "Top of Personal Folders", NID_ROOT_FOLDER, true);
            //
            // 2.7.3.4.3	Search Root
            //List<Property> searchRootReplProps = replicationItems(0x0120, 0x1f98, 0x0100);
            //Folder searchRoot = FolderPC(NID_SEARCH_FOLDER, "Search Root", NID_ROOT_FOLDER, false, searchRootReplProps);
            Folder searchRoot = FolderPC(NID_SEARCH_FOLDER, "Search Root", NID_ROOT_FOLDER, false);
            //
            // 2.7.3.4.4	Spam Search Folder
            Folder spamSearch = FolderPC(NID_SPAM_SEARCH_FOLDER, "SPAM Search Folder 2", NID_ROOT_FOLDER, false);
            //
            // 2.7.3.4.5	Deleted Items
            //List<Property> deletedReplProps = replicationItems(0x0160, 0x1fb8, 0x0140);
            Folder deletedItems = FolderPC(NID_DELETED_ITEMS, "Deleted Items", NID_TOP_PERSONAL_FOLDER, false);
            //
            //List<Property> expFolderReplProps = replicationItems(0x0120, 0x1fAC, 0x0100);
            //exportFolder.properties.AddRange(expFolderReplProps);
            //exportFolder.properties = exportFolder.properties.OrderBy(h => h.id).ToList();

            //
            // add other folder objects
            List<Folder> noSub = new List<Folder>();
            List<Folder> rootSubfolders = new List<Folder> { topPersonal, searchRoot, spamSearch };
            List<Folder> topSubfolders = new List<Folder>() { deletedItems , exportFolder};
            FolderObjects(NID_ROOT_FOLDER, rootSubfolders);
            FolderObjects(NID_TOP_PERSONAL_FOLDER, topSubfolders);
            FolderObjects(NID_SEARCH_FOLDER, noSub);
            SearchFolderObjects(NID_SPAM_SEARCH_FOLDER);
            FolderObjects (NID_DELETED_ITEMS, noSub);
        }
        private static Folder FolderPC(NID fNid, string name, NID parent, bool hasSubfolders)
        {
            List<Property> props = GetFolderPC(name, 0, 0, hasSubfolders);
            AddNIDpc(fNid, parent, props);
            return new Folder() { nid = fNid, properties = props }; 
        }
        private static void FolderObjects(NID nid, List<Folder> subFolders)
        {
            // Hierarchy TC
            NID htNid = nid.ChangeType(EnidType.HIERARCHY_TABLE);
            NID ctNid = nid.ChangeType(EnidType.CONTENTS_TABLE);
            NID asNid = nid.ChangeType(EnidType.ASSOC_CONTENTS_TABLE);
            if (subFolders.Count > 0)
            {
                TableContext TC = (TableContext)HIERARCHY_TABLE_TEMPLATE.Clone();
                TableContext tc = GetFolderHierarchyTC(subFolders);
                AddNIDtc(htNid, tc);
            }
            else
            {
                AddNIDreference(htNid, HIERARCHY_TABLE_TEMPLATE_BID, 0, 0);
            }
            AddNIDreference(ctNid, CONTENTS_TABLE_TEMPLATE_BID, 0, 0);
            AddNIDreference(asNid, ASSOC_CONTENTS_TABLE_TEMPLATE_BID, 0, 0);
        }
        private static void SearchFolderObjects(NID nid)
        {   // objects for spam search folder
            // not well doc in the MS_PST...
            NID uqNid = nid.ChangeType(EnidType.SEARCH_UPDATE_QUEUE);
            NID coNid = nid.ChangeType(EnidType.SEARCH_CRITERIA_OBJECT);
            NID ctNid = nid.ChangeType(EnidType.SEARCH_CONTENTS_TABLE);
            AddEmptyNIDentry(uqNid);
            AddSearchCriteria(coNid);
            AddNIDreference(ctNid, SEARCH_CONTENTS_TABLE_TEMPLATE_BID, 0, 0);
        }
        private static void AddSearcActivityObject()
        {
            byte[] data = new byte[4];
            FM.TypeToArray<UInt32>(ref data, NID_SPAM_SEARCH_FOLDER.dwValue);
            FM.outFile.AddNDBdata(NID_SEARCH_ACTIVITY_LIST, data,0,0);
        }
        private static void AddReceiveFolderTable()
        { // not specifed in MS_PST... replicating an empty PST view
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            TC.AddColumn((EpropertyId)0x6605, EpropertyType.PtypInteger32);
            List<Property> props = new List<Property>()
            {
                new Property(EpropertyId.PidTagMessageClassW, EpropertyType.PtypString),
                new Property((EpropertyId)0x6605, EpropertyType.PtypInteger32, NID_ROOT_FOLDER.dwValue ),
            };
            TC.AddRow(new NID(1), props);
            AddNIDtc(NID_RECIEVE_FOLDER_TABLE, TC);
        }
        private static void AddOutgoingQueueTable()
        { // not specifed in MS_PST... replicating an empty PST view
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagClientSubmitTime, EpropertyType.PtypTime);
            TC.AddColumn((EpropertyId)0x0E10, EpropertyType.PtypInteger32);
            TC.AddColumn((EpropertyId)0x0E14, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            AddNIDtc(NID_OUTGOING_QUEUE_TABLE, TC);
        }
        // the NID 3073 is not documented in MS_PST, but required by outlook/SCANPST
        // I have just copied the data from an empty PST file
        // outlook accepts it... SCANPST does not flag as an error but generates
        // a new one during repair....
        private static readonly byte[] nid3073data = new byte[34] { // from a empty pst
            0x18, 0x00, 0xEC, 0x9C, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00,
            0xB5, 0x10, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x0C, 0x00, 0x10, 0x00,
            0x18, 0x00
        };
        private static void AddUndocObjects()
        {
            //undocTable1();
            undocTable2();
            undocTable3();
            FM.outFile.AddNDBdata(new NID(0x01E1), 0, 0, 0);
           // FM.outFile.AddNDBdata(new NID(0x0261), 0, 0, 0);
            FM.outFile.AddNDBdata(new NID(0x0EC1), 0, 0, 0);
            FM.outFile.AddNDBdata(new NID(0x0C01), nid3073data, 0, 0);
        }
        private static void undocTable1()
        {   // not specifed in MS_PST... replicating an empty PST view
            TableContext TC = new TableContext();
            TC.AddColumn((EpropertyId)0x0E33, EpropertyType.PtypInteger64);
            TC.AddColumn((EpropertyId)0x0E37, (EpropertyType)0x102);
            TC.AddColumn((EpropertyId)0x0E38, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            AddNIDtc(new NID(0x06B6), TC);
        }
        private static void undocTable2()
        {   // not specifed in MS_PST... replicating an empty PST view
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            TC.AddColumn((EpropertyId)0x0E30, (EpropertyType)0x102);
            TC.AddColumn((EpropertyId)0x0E31, (EpropertyType)0x102);
            TC.AddColumn((EpropertyId)0x0E33, (EpropertyType)0x14);
            TC.AddColumn((EpropertyId)0x0E34, (EpropertyType)0x102);
            TC.AddColumn((EpropertyId)0x0E38, (EpropertyType)0x3);
            TC.AddColumn((EpropertyId)0x0E3E, (EpropertyType)0x102);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            AddNIDtc(new NID(0x06D7), TC);
        }
        private static void undocTable3()
        {   // not specifed in MS_PST... replicating an empty PST view
            TableContext TC = new TableContext();
            TC.AddColumn((EpropertyId)0x0E33, (EpropertyType)0x14);
            TC.AddColumn((EpropertyId)0x3007, (EpropertyType)0x40);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            AddNIDtc(new NID(0x06F8), TC);
        }
        private static TableContext GetFolderHierarchyTC(List<Folder> subFolders)
        {
            TableContext TC = (TableContext)HIERARCHY_TABLE_TEMPLATE.Clone();
            TC.resetRows();
            subFolders = subFolders.OrderBy(s => s.nid.dwValue).ToList();
            foreach (Folder sf in subFolders)
            {
                List<Property> props = new List<Property>();
                props.AddRange(sf.properties);
                if (sf.nid.dwValue != 8739)
                {   // SPAM Search Folder 2 seems not to have repl items
                    props.AddRange(LTP.ReplTagProps());
                }
                // TC.AddRow(sf.nid, sf.properties);
                TC.AddRow(sf.nid, props);
            }
            return TC;
        }
        public static void AddSearchCriteria(NID nid)
        {   // not documented... just checked a pst file
            List<Property> props = new List<Property>() { new Property((EpropertyId)(0x660B), EpropertyType.PtypInteger32, 0) };
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            FM.outFile.AddNDBdata(nid, pcHNblocks, 0, 0);
        }
        public static void AddNIDpc(NID nid, NID parent, List<Property> props)
        {   // just for top folders... no subnode
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            UInt64 newBid = FM.outFile.AddNDBdata(nid, pcHNblocks, 0, parent.dwValue);
        }
        public static UInt64 AddNIDtc(NID nid, TableContext tc)
        {   // just for template tc's... no subnode
            BTH tcBTHs = new BTH(tc);
            List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
            return FM.outFile.AddNDBdata(nid, tcHNblocks, 0, 0);
        }
        public static void ValidateFolderPC(List<Property> pc)
        {   // MS_PST 2.4.4.1.1	Property Schema of a Folder object PC
            CheckRequiredProperty(pc, EpropertyId.PidTagDisplayName, EpropertyType.PtypString);
            CheckRequiredProperty(pc, EpropertyId.PidTagContentCount, EpropertyType.PtypInteger32);
            CheckRequiredProperty(pc, EpropertyId.PidTagContentUnreadCount, EpropertyType.PtypInteger32);
            CheckRequiredProperty(pc, EpropertyId.PidTagSubfolders, EpropertyType.PtypBoolean);
            RemoveProperty(pc, (EpropertyId)0x67F4);

        }
        private static void RemoveProperty(List<Property> props, EpropertyId id)
        {
            int propIx = props.FindIndex(f => f.id == id);
            if (propIx >= 0)
            {
                props.RemoveAt(propIx);
            }
        }
        public static void ValidateMessagePC(List<Property> pc)
        {   // MS_PST 2.4.5.1.1	Property Schema of a Folder object PC
            CheckRequiredProperty(pc, EpropertyId.PidTagMessageClass, EpropertyType.PtypString);
            CheckRequiredProperty(pc, EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(pc, EpropertyId.PidTagMessageSize, EpropertyType.PtypInteger32);
            CheckRequiredProperty(pc, EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            CheckRequiredProperty(pc, EpropertyId.PidTagCreationTime, EpropertyType.PtypTime);
            CheckRequiredProperty(pc, EpropertyId.PidTagLastModificationTime, EpropertyType.PtypTime);
            CheckRequiredProperty(pc, EpropertyId.PidTagSearchKey, EpropertyType.PtypBinary);
            RemoveProperty(pc, (EpropertyId)0x67F4);
        }
        public static void ValidateContentsTable(ref TableContext tc)
        {   // MS_PST 2.4.4.5.1	Contents Table Template
            CheckRequiredProperty(ref tc, EpropertyId.PidTagImportance, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSensitivity, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSubjectW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagClientSubmitTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSentRepresentingNameW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageToMe, EpropertyType.PtypBoolean);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageCcMe, EpropertyType.PtypBoolean);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagConversationTopicW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagConversationIndex, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagDisplayCcW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagDisplayToW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageDeliveryTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageSize, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            //CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypInteger32);  // scanpst expected binary type
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplCopiedfromVersionhistory, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplCopiedfromItemid, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagItemTemporaryFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLastModificationTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSecureSubmitFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            tc.RemoveColumn((EpropertyId)0x67f4, EpropertyType.PtypInteger64);
        }
        public static void ValidateAssocContentsTable(ref TableContext tc)
        {   // MS_PST 2.4.4.5.1	Contents Table Template
            CheckRequiredProperty(ref tc, EpropertyId.PidTagImportance, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSensitivity, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSubjectW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagClientSubmitTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSentRepresentingNameW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageToMe, EpropertyType.PtypBoolean);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageCcMe, EpropertyType.PtypBoolean);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagConversationTopicW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagConversationIndex, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagDisplayCcW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagDisplayToW, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageDeliveryTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageSize, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            //CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypInteger32);  // scanpst expected binary type
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplCopiedfromVersionhistory, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplCopiedfromItemid, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagItemTemporaryFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLastModificationTime, EpropertyType.PtypTime);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSecureSubmitFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            tc.RemoveColumn((EpropertyId)0x67f4, EpropertyType.PtypInteger64);
        }
        public static void ValidateHierarchyTable(ref TableContext tc)
        {   // MS_PST 2.4.4.4.1	Hierarchy Table Template
            //CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypInteger32);  // scanpst expected binary type
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagDisplayName, EpropertyType.PtypString);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagContentCount, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagContentUnreadCount, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagSubfolders, EpropertyType.PtypBoolean);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagContainerClass, EpropertyType.PtypBinary);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagPstHiddenCount, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagPstHiddenUnread, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            CheckRequiredProperty(ref tc, EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            tc.RemoveColumn((EpropertyId)0x67f4, EpropertyType.PtypInteger64);
        }
        private static void CheckRequiredProperty(ref TableContext tc, EpropertyId id, EpropertyType type)
        {
            if (tc.FindColumn(id, type) < 0)
            {
                tc.AddColumn(id, type);
            }
        }
        private static void CheckRequiredProperty(List<Property> pc, EpropertyId id, EpropertyType type)
        {
            if (pc.FindIndex(p => (p.PCBTH.wPropId == id & p.PCBTH.wPropType == type)) < 0)
            {
                Property prop = new Property(id, type);
                pc.Add(prop);
            }
        }

        private static TableContext CREATE_HIERARCHY_TABLE_TEMPLATE()
        {   // MS_PST 2.4.4.4.1	Hierarchy Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary);          // 2.4.4.4.1 states int32 but outlook bynary
            TC.AddColumn(EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64);
            TC.AddColumn(EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagDisplayName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagContentCount, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagContentUnreadCount, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagSubfolders, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagContainerClass, EpropertyType.PtypString);      // 2.4.4.4.1 states binary but outlook string
            TC.AddColumn(EpropertyId.PidTagPstHiddenCount, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagPstHiddenUnread, EpropertyType.PtypInteger32);
            HIERARCHY_TABLE_TEMPLATE = TC;
            return TC;
        }
        public static TableContext CREATE_CONTENTS_TABLE_TEMPLATE()
        {   // MS_PST 2.4.4.5.1	Contents Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagImportance, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSensitivity, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagSubjectW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagClientSubmitTime, EpropertyType.PtypTime);
            TC.AddColumn(EpropertyId.PidTagSentRepresentingNameW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagMessageToMe, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagMessageCcMe, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagConversationTopicW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagConversationIndex, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagDisplayCcW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagDisplayToW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagMessageDeliveryTime, EpropertyType.PtypTime);
            TC.AddColumn(EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageSize, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary);           // 2.4.4.4.1 states int32 but outlook bynary
            TC.AddColumn(EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64);
            TC.AddColumn(EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagReplFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagReplCopiedfromVersionhistory, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagReplCopiedfromItemid, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagItemTemporaryFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLastModificationTime, EpropertyType.PtypTime);
            TC.AddColumn(EpropertyId.PidTagSecureSubmitFlags, EpropertyType.PtypInteger32);
            TC.AddColumn((EpropertyId)0x3013, EpropertyType.PtypBinary);                        // flaged as missing in scanpst 
            //            CONTENTS_TABLE_TEMPLATE = (TableContext)TC.Clone();
            // AddNIDtc(NID.contents_table_temp, TC);
            CONTENTS_TABLE_TEMPLATE = TC;
            return TC;
        }
        public static TableContext CREATE_ASSOC_CONTENTS_TABLE_TEMPLATE()
        {   // MS_PST 2.4.4.6.1	FAI Contents Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagMessageClass, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagDisplayName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagOfflineAddressBookName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSendOutlookRecallReport, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagOfflineAddressBookTruncatedProperties, EpropertyType.PtypMultipleInteger32);
            TC.AddColumn(EpropertyId.PidTagViewDescriptorFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagViewDescriptorLinkTo, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagViewDescriptorViewFolder, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagViewDescriptorName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagViewDescriptorVersion, EpropertyType.PtypInteger32);
            TC.AddColumn((EpropertyId)0x682f, EpropertyType.PtypString);                        // flaged as missing in scanpst 
            //  ASSOC_CONTENTS_TABLE_TEMPLATE = (TableContext)TC.Clone();
            // AddNIDtc(NID.assoc_contents_table_temp, TC);
            ASSOC_CONTENTS_TABLE_TEMPLATE = TC;
            return TC;
        }
        public static TableContext CREATE_SEARCH_CONTENTS_TABLE_TEMPLATE()
        {   // MS_PST 2.4.8.6.2.1	Search Folder Contents Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagImportance, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageClassW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSensitivity, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageFlags, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagMessageStatus, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagSubjectW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSentRepresentingNameW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagMessageToMe, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagMessageCcMe, EpropertyType.PtypBoolean);      // scanpst requires this coll... not in MS_PST
            TC.AddColumn(EpropertyId.PidTagDisplayCcW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagDisplayToW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagParentDisplayW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagMessageDeliveryTime, EpropertyType.PtypTime);
            TC.AddColumn(EpropertyId.PidTagMessageSize, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagExchangeRemoteHeader, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagLastModificationTime, EpropertyType.PtypTime);
            TC.AddColumn(EpropertyId.PidTagLtpParentNid, EpropertyType.PtypInteger32);
            // AddNIDtc(NID.search_contents_table_temp, TC);
            SEARCH_CONTENTS_TABLE_TEMPLATE = TC;
            return TC;
        }
        public static TableContext CREATE_RECIPIENT_TABLE_TEMPLATE()
        {   // MS_PST 2.4.5.3.1	Recipient Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagRecipientType, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagResponsibility, EpropertyType.PtypBoolean);
            TC.AddColumn(EpropertyId.PidTagRecordKey, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagObjectType, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagEntryId, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagDisplayName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagAddressType, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagEmailAddress, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSearchKey, EpropertyType.PtypBinary);
            TC.AddColumn(EpropertyId.PidTagDisplayType, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTag7BitDisplayName, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagSendRichInfo, EpropertyType.PtypBoolean);
            // AddNIDtc(NID.recipient_table_temp, TC);
            RECIPIENT_TABLE_TEMPLATE = TC;
            return TC;
        }
        public static TableContext CREATE_ATTACHMENT_TABLE_TEMPLATE()
        {   // MS_PST 2.4.6.1.1	Attachment Table Template
            TableContext TC = new TableContext();
            TC.AddColumn(EpropertyId.PidTagAttachSize, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagAttachFilenameW, EpropertyType.PtypString);
            TC.AddColumn(EpropertyId.PidTagAttachMethod, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagRenderingPosition, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowId, EpropertyType.PtypInteger32);
            TC.AddColumn(EpropertyId.PidTagLtpRowVer, EpropertyType.PtypInteger32);
            // AddNIDtc(NID.attachment_table_tep, TC);
            ATTACHMENT_TABLE_TEMPLATE = TC;
            return TC;
        }
        private static byte[] NidToArray(NID nid)
        {
            byte[] nidArray = new byte[4];
            FM.TypeToArray<UInt32>(ref nidArray, nid.dwValue);
            return nidArray;
        }
        private static Property SetPropertyValue(EpropertyId propId, EpropertyType propType, dynamic value)
        {
            Property PC = new Property(propId, propType, value);
            switch (propType)
            {
                case EpropertyType.PtypInteger16:
                    PC.setHNID(new HNID((UInt16)value));
                    break;

                case EpropertyType.PtypInteger32:
                    PC.setHNID(new HNID((UInt32)value));
                    break;

                case EpropertyType.PtypBoolean:
                    uint val = value ? 1u : 0u;
                    PC.setHNID(new HNID(val));
                    break;
            }
            return PC;
        }

        private static void AddNIDreference(NID nid, UInt64 bid, UInt32 nidParent, UInt64 bidSub)
        {
            FM.outFile.AddNDBdata(nid, bid, nidParent, bidSub);
            FM.outFile.IncrementRefCount(bid);
        }
        private static void AddEmptyNIDentry(NID nid)
        {
            FM.outFile.AddNDBdata(nid, 0, 0, 0);
        }

        public static List<Property> GetFolderPC(string name, uint contentCount, uint contentUnreadCount, bool subfolders)
        {
            List<Property> props = new List<Property>();
            props.Add(new Property(EpropertyId.PidTagDisplayName, EpropertyType.PtypString, name));
            props.Add(new Property(EpropertyId.PidTagContentCount, EpropertyType.PtypInteger32, contentCount));
            props.Add(new Property(EpropertyId.PidTagContentUnreadCount, EpropertyType.PtypInteger32, contentUnreadCount));
            props.Add(new Property(EpropertyId.PidTagSubfolders, EpropertyType.PtypBoolean, subfolders));
            return props;
        }
    }
}