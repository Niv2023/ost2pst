using System.Data;
using System.Runtime.InteropServices;
using System.Text;

namespace ost2pst
{
    public class Folder
    {
        public int nbtIndex;
        public int level;
        public Folder parent;
        public string name;
        public string path;
        public NID nid;
        public bool toBeExported;
        public List<Property> properties;
    }
    public class msgSize
    {
        public UInt32 nid;
        public UInt32 size;
    }

    public static class FM
    {
        public static XstFile srcFile;
        public static PstFile outFile;
        public static List<msgSize> msgSizes;
        public static List<Folder> folders;
        public static Folder RootFolder;
        public static UInt64 defaulTime;
        public static bool cleanBID = true;
        public static bool OpenSourceFile(string filename)
        {
            defaulTime = (UInt64)DateTime.Now.ToFileTimeUtc();   // used when required timestamps are missing
            bool result = false;
            try
            {
                srcFile = new XstFile(filename);
                result = true;
            }
            catch (Exception ex)
            {
                Program.mainForm.statusMSG(ex.Message);
            }
            return result;
        }
        public static bool CreatPstFile(string filename)
        {
            bool result = false;
            try
            {
                outFile = new PstFile(filename);
                result = true;
            }
            catch (Exception ex)
            {
                Program.mainForm.statusMSG(ex.Message);
            }
            return result;
        }
        public static void RebuildMessageObject(NBTENTRY nbt)
        {   // MS_PST 2.4.5	Message Objects
            List<SLENTRY> newSubnodes = new();
            List<SLENTRY> msgSubnodes = FM.srcFile.GetSLentries(nbt.bidSub);
            foreach (SLENTRY sn in msgSubnodes)
            {
                SLENTRY newSn = new SLENTRY()
                {
                    nid = sn.nid,
                    bidSub = 0
                };
                UpdateSubnodeNIDheaderIndex((UInt32)sn.nid);
                switch (NID.Type(sn.nid))
                {
                    case EnidType.RECIPIENT_TABLE:          // mandatory
                        newSn = RebuildMessageRecipientTC(sn);
                        break;
                    case EnidType.ATTACHMENT_TABLE:         // optional
                        newSn = RebuildMessageRecipientTC(sn);
                        break;
                    case (EnidType)0x16:  // undoc nid that is present in msg objects it is a TC
                        newSn = RebuildMessageRecipientTCrefreshIndex(sn);
                        break;
                    case EnidType.ATTACHMENT:               // optional
                        List<Property> props = LTP.ReadSubnodePC(srcFile.stream, sn);
                        int attObjectIx = props.FindIndex(p => p.id == EpropertyId.PidTagAttachDataBinary & p.type == EpropertyType.PtypObject);
                        if (attObjectIx >= 0)
                        {   // attachment is a message object see MS_PST:
                            // 2.3.3.5	PtypObject Properties
                            // 2.4.6.2.2	Attachment Data
                            UInt32 objNid = ExtractTypeFromArray<UInt32>(props[attObjectIx].data);
                            UInt32 objSize = ExtractTypeFromArray<UInt32>(props[attObjectIx].data, 4);
                            if ((EnidType)(objNid & 0x1f) == EnidType.NORMAL_MESSAGE)
                            {
                                List<SLENTRY> subnodes = FM.srcFile.GetSLentries(sn.bidSub);
                                int msgSnIx = subnodes.FindIndex(s => s.nid == objNid);
                                if (msgSnIx < 0) throw new Exception($"Attachment msg obj {objNid} subnode not found");
                                SLENTRY msgSL = subnodes[msgSnIx];
                                NBTENTRY msgNBT = new NBTENTRY()
                                {
                                    nid = new NID(objNid),
                                    bidData = msgSL.bidData,
                                    bidSub = msgSL.bidSub,
                                    nidParent = 0,
                                };
                                UpdateSubnodeNIDheaderIndex(msgNBT.nid.dwValue);
                                NBTENTRY newNbt = RebuildAttachedMessageObject(msgNBT);
                                msgSL.bidData = newNbt.bidData;
                                msgSL.bidSub = newNbt.bidSub;
                                newSn = RebuildAttachmnetPC(objNid, props, subnodes, msgSL);
                                newSn.nid = sn.nid;
                            }
                            else
                            {
                                newSn = RebuildAttachmnetPC(sn);
                            }
                        }
                        else
                        {
                            newSn = RebuildAttachmnetPC(sn);
                        }
                        break;
                    case EnidType.LTP:
                        newSn = Blocks.CopySLentry(sn);
                        break;
                    default:

                        /* MS_PST 2.4.5	Message Objects 
                         * other nid type not specified
                         */
                        throw new Exception($"Message ({nbt.nid.dwValue}) with an invalid subnode nid type {sn.nid}");
                }
                newSubnodes.Add(newSn);
            }
            BREF snBref = Blocks.AddSLentries(newSubnodes);
            RebuildMessagePC(nbt, snBref.bid);
        }
        public static NBTENTRY RebuildAttachedMessageObject(NBTENTRY nbt)
        {   // MS_PST 2.4.5	Message Objects
            List<SLENTRY> newSubnodes = new();
            List<SLENTRY> msgSubnodes = FM.srcFile.GetSLentries(nbt.bidSub);
            foreach (SLENTRY sn in msgSubnodes)
            {
                SLENTRY newSn = new SLENTRY()
                {
                    nid = sn.nid,
                    bidSub = 0
                };
                UpdateSubnodeNIDheaderIndex((UInt32)sn.nid);
                switch (NID.Type(sn.nid))
                {
                    case EnidType.RECIPIENT_TABLE:          // mandatory
                        newSn = RebuildMessageRecipientTC(sn);
                        break;
                    case EnidType.ATTACHMENT_TABLE:         // optional
                        newSn = RebuildAttachmnetTableTC(sn);
                        break;
                    case EnidType.ATTACHMENT:               // optional
                        List<Property> props = LTP.ReadSubnodePC(srcFile.stream, sn);
                        int attObjectIx = props.FindIndex(p => p.id == EpropertyId.PidTagAttachDataBinary & p.type == EpropertyType.PtypObject);
                        if (attObjectIx >= 0)
                        {   // attachment is a message object see MS_PST:
                            // 2.3.3.5	PtypObject Properties
                            // 2.4.6.2.2	Attachment Data
                            UInt32 objNid = ExtractTypeFromArray<UInt32>(props[attObjectIx].data);
                            UInt32 objSize = ExtractTypeFromArray<UInt32>(props[attObjectIx].data, 4);
                            if ((EnidType)(objNid & 0x1f) == EnidType.NORMAL_MESSAGE)
                            {
                                List<SLENTRY> subnodes = FM.srcFile.GetSLentries(sn.bidSub);
                                int msgSnIx = subnodes.FindIndex(s => s.nid == objNid);
                                if (msgSnIx < 0) throw new Exception($"Attachment msg obj {objNid} subnode not found");
                                SLENTRY msgSL = subnodes[msgSnIx];
                                NBTENTRY msgNBT = new NBTENTRY()
                                {
                                    nid = new NID(objNid),
                                    bidData = msgSL.bidData,
                                    bidSub = msgSL.bidSub,
                                    nidParent = 0,
                                };
                                UpdateSubnodeNIDheaderIndex(msgNBT.nid.dwValue);
                                NBTENTRY newNbt = RebuildAttachedMessageObject(msgNBT);
                                msgSL.bidData = newNbt.bidData;
                                msgSL.bidSub = newNbt.bidSub;
                                newSn = RebuildAttachmnetPC(objNid, props, subnodes, msgSL);
                                newSn.nid = sn.nid;
                            }
                            else
                            {
                                newSn = RebuildAttachmnetPC(sn);
                            }
                        }
                        else
                        {
                            newSn = RebuildAttachmnetPC(sn);
                        }
                        break;
                    default:
                        newSn = Blocks.CopySLentry(sn);
                        break;
                }
                newSubnodes.Add(newSn);
            }
            BREF snBref = Blocks.AddSLentries(newSubnodes);
            return RebuildSubnodeMessagePC(nbt, snBref.bid);
        }
        private static SLENTRY RebuildAttachmnetPC(SLENTRY sn)
        {
            UInt64 subnode = 0;
            if (sn.bidSub > 0)
            {   
                BREF snBref = Blocks.CopySubnode(sn.bidSub);
                subnode = snBref.bid;
            }
            List<Property> props = LTP.ReadSubnodePC(srcFile.stream, sn);
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            return new SLENTRY()
            {
                nid = sn.nid,
                bidData = AddSubnode(pcHNblocks, sn.bidData),
                bidSub = subnode
            };
        }
        private static SLENTRY RebuildAttachmnetPC(UInt64 nid, List<Property> props, List<SLENTRY> sn, SLENTRY msgSL)
        {
            BREF slBref = Blocks.CopySubnodes(sn, msgSL);
            UInt64 subnode = slBref.bid;
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            return new SLENTRY()
            {
                nid = nid,
                bidData = AddSubnode(pcHNblocks),
                bidSub = subnode
            };
        }
        private static SLENTRY RebuildAttachmnetTableTC(SLENTRY sn)
        {
            SLENTRY slTC = new SLENTRY() { nid = sn.nid };
            TableContext tc = LTP.ReadSubnodeTC(srcFile.stream, sn);  // validate????
            BTH tcBTHs = new BTH(tc, true);
            if (tcBTHs.subnodes.Count > 0)
            {
                slTC.bidSub = Blocks.AddSLentries(tcBTHs.subnodes).bid;
            };
            List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
            slTC.bidData = AddSubnode(tcHNblocks);
            return slTC;
        }
        private static SLENTRY RebuildMessageRecipientTC(SLENTRY sn)
        {
            UInt64 subnode = 0;
            TableContext tc = LTP.ReadSubnodeTC(srcFile.stream, sn);
            BTH tcBTHs = new BTH(tc);
            if (sn.bidSub > 0)
            {   
                List<byte[]> data = new();
                List<SLENTRY> pstSN = new List<SLENTRY>();
                List<SLENTRY> ostSN = FM.srcFile.GetSLentries(sn.bidSub);
                foreach (SLENTRY ostEntry in ostSN)
                {
                    if (ostEntry.nid == tc.tcINFO.hnidRows.dwValue)
                    {   // rebuilt TC data rows
                        data = tc.RowMatrixArray();
                    }
                    else
                    {
                        data = new List<byte[]> { FM.srcFile.ReadFullDatablock(ostEntry.bidData) };
                    }
                    SLENTRY sl = new SLENTRY()
                    {
                        nid = ostEntry.nid,
                        bidData = AddSubnode(data, ostEntry.bidData),
                        bidSub = ostEntry.bidSub
                    };
                    UpdateSubnodeNIDheaderIndex((UInt32)ostEntry.nid);
                    pstSN.Add(sl);
                }
                BREF snBref = Blocks.AddSLentries(pstSN);
                subnode = snBref.bid;
            }
            List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
            return new SLENTRY()
            {
                nid = sn.nid,
                bidData = AddSubnode(tcHNblocks, sn.bidData),
                bidSub = subnode
            };
        }
        private static SLENTRY RebuildMessageRecipientTCrefreshIndex(SLENTRY sn)
        {
            UInt64 subnode = 0;
            TableContext tc = LTP.ReadSubnodeTCandRefreshIndex(srcFile.stream, sn);
            BTH tcBTHs = new BTH(tc);
            if (sn.bidSub > 0)
            {   
                List<byte[]> data = new();
                List<SLENTRY> pstSN = new List<SLENTRY>();
                List<SLENTRY> ostSN = FM.srcFile.GetSLentries(sn.bidSub);
                foreach (SLENTRY ostEntry in ostSN)
                {
                    if (ostEntry.nid == tc.tcINFO.hnidRows.dwValue)
                    {   // rebuilt TC data rows
                        data = tc.RowMatrixArray();
                    }
                    else
                    {
                        data = new List<byte[]> { FM.srcFile.ReadFullDatablock(ostEntry.bidData) };
                    }
                    SLENTRY sl = new SLENTRY()
                    {
                        nid = ostEntry.nid,
                        bidData = AddSubnode(data, ostEntry.bidData),
                        bidSub = ostEntry.bidSub
                    };
                    UpdateSubnodeNIDheaderIndex((UInt32)ostEntry.nid);
                    pstSN.Add(sl);
                }
                BREF snBref = Blocks.AddSLentries(pstSN);
                subnode = snBref.bid;
            }
            List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
            return new SLENTRY()
            {
                nid = sn.nid,
                bidData = AddSubnode(tcHNblocks, sn.bidData),
                bidSub = subnode
            };
        }

        private static UInt64 AddSubnode(List<byte[]> data, UInt64 srcBid = 0)
        {
            BREF bref = Blocks.AddDatablock(data, srcBid);
            return bref.bid;
        }
        public static NBTENTRY RebuildMessagePC(NBTENTRY nbt, UInt64 newSubnode)
        {
            List<Property> props = ReadAndValidateOSTpcs(nbt);
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            NBTENTRY newNbt = outFile.AddNDBdataNbt(nbt.nid, pcHNblocks, newSubnode, nbt.nidParent);
            return newNbt;
        }
        public static NBTENTRY RebuildSubnodeMessagePC(NBTENTRY nbt, UInt64 newSubnode)
        {
            List<Property> props = ReadAndValidateOSTpcs(nbt);
            BTH pcBTHs = new BTH(props);
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            NBTENTRY newNbt = outFile.AddSubnodeNDBdataNbt(nbt.nid, pcHNblocks, newSubnode, nbt.nidParent);
            return newNbt;
        }
        public static void ResetPasswordPC(NBTENTRY nbt)
        {
            List<Property> props = LTP.ReadPCs(srcFile.stream, nbt);
            int ixPW = props.FindIndex(p => p.id == EpropertyId.PidTagPstPassword);
            if (ixPW < 0) return; // there is no password set in the message store
            PCBTH pc = new PCBTH()
            {
                wPropId = EpropertyId.PidTagPstPassword,
                wPropType = EpropertyType.PtypInteger32,
                dwValueHnid = new HNID(0)
            };
            props[ixPW] = new(pc);
            BTH pcBTHs = new BTH(props);  
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            outFile.AddNDB(nbt, pcHNblocks, 0);  // no subnode for PC on a message store
        }
        public static void RefreshMessagePC(NBTENTRY nbt)
        {
            List<Property> props = LTP.ReadPCs(srcFile.stream,nbt);
            UInt64 subnode = 0;
            UInt32 ms = 0;
            int pSix = props.FindIndex(p => p.id == EpropertyId.PidTagMessageSize);
            msgSize? mSize = msgSizes.Find(m => m.nid == nbt.nid.dwValue);
            if (mSize == null || pSix <0) { throw new Exception($"missing message size nid {nbt.nid.dwValue}"); }
            Property prop = props[pSix];
            ms = ExtractTypeFromArray<UInt32>(prop.data);
            if (mSize.size != ms)
            {
                PCBTH pc = new PCBTH()
                {
                    wPropId = prop.id,
                    wPropType = prop.type,
                    dwValueHnid = new HNID(mSize.size)
                };
                props[pSix] = new(pc);
                BTH pcBTHs = new BTH(props);  // props with subnode ref remains unchanged
                if (nbt.bidSub != 0)
                {
                    BREF snBref = Blocks.CopySubnode(nbt.bidSub);
                    subnode = snBref.bid;
                }
                List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
                outFile.AddNDB(nbt, pcHNblocks, subnode);
            }
            else
            {    // msg PC has the correct size... just copy the nbd
                outFile.CopyNDB(nbt);
            }
        }
        private static bool MsgSizeChanged(TableContext tc)
        {
            if (tc.tcRowIndexes.Count == 0) return false;
            bool hasChanged = false;
            for (int i = 0; i < tc.tcRowIndexes.Count; i++)
            {
                msgSize? mSize = msgSizes.Find(m => m.nid == tc.tcRowIndexes[i].dwRowID);
                if (mSize == null ) continue;   // row id is not message nid
                RowData rd = tc.tcRowMatrix[i];
                int pSix = rd.Props.FindIndex(p => p.id == EpropertyId.PidTagMessageSize);
                Property prop = rd.Props[pSix];
                UInt32 ms = ExtractTypeFromArray<UInt32>(prop.data);
                if (mSize.size != ms)
                {
                    hasChanged = true;
                    PCBTH pc = new PCBTH()
                    {
                        wPropId = prop.id,
                        wPropType = prop.type,
                        dwValueHnid = new HNID(mSize.size)
                    };
                    rd.Props[pSix] = new(pc);
                }
            }   
            return hasChanged;
        }

        public static void RefreshContentTableTC(NBTENTRY nbt)
        {   
            TableContext tc = LTP.ReadTCs_and_rowdata(srcFile.stream, nbt);
            if (MsgSizeChanged(tc))
            {
                BTH tcBTHs = new BTH(tc, true);
                UInt64 subnode = 0;
                if (tcBTHs.subnodes.Count > 0)
                {  
                    BREF snBREF = Blocks.AddSLentries(tcBTHs.subnodes);
                    subnode = snBREF.bid;
                };
                List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
                outFile.AddNDB(nbt, tcHNblocks, subnode);
            }
            else
            {  // 
                outFile.CopyNDB(nbt);
            }
        }
        public static void RebuildPC(NBTENTRY nbt)
        {
            List<Property> props = ReadAndValidateOSTpcs(nbt);
            UInt64 subnode = 0;
            BTH pcBTHs = new BTH(props);  // props with subnode ref remains unchanged
            if (nbt.bidSub != 0)
            {
                BREF snBref = Blocks.CopySubnode(nbt.bidSub);
                subnode = snBref.bid;
            }
            List<byte[]> pcHNblocks = LTP.GetBTHhnDatablocks(pcBTHs, EbType.bTypePC);
            outFile.AddNDB(nbt, pcHNblocks, subnode);
        }
        public static void RebuildTC(NBTENTRY nbt)
        {   // rebuild TC
            // row ids are considered NID PCs
            // the PCs are read for reconstructing the row matrix
            TableContext tc = ReadAndValidateOSTtcs(nbt);
            BTH tcBTHs = new BTH(tc, true);
            UInt64 subnode = 0;
            if (tcBTHs.subnodes.Count > 0)
            {  
                BREF snBREF = Blocks.AddSLentries(tcBTHs.subnodes);
                subnode = snBREF.bid;
            };
            List<byte[]> tcHNblocks = LTP.GetBTHhnDatablocks(tcBTHs, EbType.bTypeTC);
            outFile.AddNDB(nbt, tcHNblocks, subnode);
        }
        private static List<Property> ReadAndValidateOSTpcs(NBTENTRY nbt)
        {
            List<Property> props = LTP.ReadPCs(srcFile.stream, nbt);
            if (nbt.nid.nidType == EnidType.NORMAL_FOLDER) RQS.ValidateFolderPC(props);
            else if (nbt.nid.nidType == EnidType.NORMAL_MESSAGE || nbt.nid.nidType == EnidType.ASSOC_MESSAGE) RQS.ValidateMessagePC(props);
            props = props.OrderBy(h => h.PCBTH.wPropId).ToList();
            return props;
        }
        private static TableContext ReadAndValidateOSTtcs(NBTENTRY nbt)
        {
            TableContext tc = LTP.ReadTCs(srcFile.stream, nbt);
            if (nbt.nid.nidType == EnidType.CONTENTS_TABLE) RQS.ValidateContentsTable(ref tc);
            else if (nbt.nid.nidType == EnidType.HIERARCHY_TABLE) RQS.ValidateHierarchyTable(ref tc);
            else if (nbt.nid.nidType == EnidType.ASSOC_CONTENTS_TABLE) RQS.ValidateAssocContentsTable(ref tc);
            return tc;
        }
        private static UInt32 GetActualMessageSize(NBTENTRY nbt)
        {
            if (nbt.nid.nidType != EnidType.NORMAL_MESSAGE) throw new Exception($"nid {nbt.nid.dwValue} is not a message nid");
            UInt32 msgSize = bidDataSize(nbt.bidData);
            if (nbt.bidSub > 0) msgSize += bidSubDataSize(nbt.bidSub);
            return msgSize;
        }
        private static UInt32 bidSubDataSize(UInt64 bidSub)
        {
            UInt32 bidSubSize = 0;
            List<SLENTRY> srcSubnodes = FM.srcFile.GetSLentries(bidSub);
            foreach (SLENTRY s in srcSubnodes)
            {
                bidSubSize += bidDataSize(s.bidData);
                if (s.bidSub >0) bidSubSize += bidSubDataSize(s.bidSub);
            }
            return bidSubSize;
        }
        private static UInt32 bidDataSize(UInt64 bid)
        {
            BBT msgDataBBT = srcFile.BBTs.Find(b => b.BREF.bid == bid);
            if (msgDataBBT.type == EbType.bTypeD)
            {
                return msgDataBBT.cbInflated;
            }
            else
            {
                return msgDataBBT.IcbTotal;
            }
        }
        public static void rebuildPSTfile(bool resetPassword)
        {
            getMsgSizes();
            CopyRevisedPSTdata(resetPassword);
        }
        private static void getMsgSizes()
        {
            msgSizes = new();
            List<NBTENTRY> msgNbts = (List<NBTENTRY>)srcFile.NBTs.FindAll(n => n.nid.nidType == EnidType.NORMAL_MESSAGE);
            Program.mainForm.statusMSG($"Reading pst message sizes");
            foreach (NBTENTRY nbt in msgNbts)
            {
                msgSizes.Add(new()
                {
                    nid = nbt.nid.dwValue,
                    size = GetActualMessageSize(nbt)
                });
            }
        }
        public static void CopyRevisedPSTdata(bool resetPassword)
        {
            Program.mainForm.statusMSG($"Exporting NBT data");
            for (int i = 0; i < srcFile.NBTs.Count; i++)
            {
                NBTENTRY nbt = srcFile.NBTs[i];
                Program.mainForm.statusMSG($"Exporting NBT entry {i+1} out of {srcFile.NBTs.Count}",false);
                if (nbt.nid.nidType == EnidType.NORMAL_MESSAGE)
                {
                   RefreshMessagePC(nbt);
                }
                else
                {
                   if (nbt.nid.nidType == EnidType.CONTENTS_TABLE)
                   {
                       RefreshContentTableTC(nbt);
                   }
                   else
                   {
                        if (nbt.bidData == 0)
                        {
                            outFile.NBTs.Add(nbt); // nbt without data just copy the entry
                        }
                        else
                        {
                            if (resetPassword && nbt.nid.dwValue == (uint)EnidSpecial.NID_MESSAGE_STORE)
                            {   // reset password for message store
                                ResetPasswordPC(nbt);
                            }
                            else
                            {
                                outFile.CopyNDB(nbt);
                            }
                        }
                    }
                }
            }
        }

        public static void CopySourceDatablocksToPST(string folderToExport, string filename)
        {
            Folder expFolder = CheckFoldersToExport(folderToExport);
            RQS.BuildPSTobjects(expFolder, filename);
            for (int i = 0; i < srcFile.NBTs.Count; i++)
            {
                NBTENTRY nbt = srcFile.NBTs[i];
                Program.mainForm.statusMSG($"Converting OST NBT entry {i + 1} out of {srcFile.NBTs.Count}", false);
                if (ToBeExported(nbt))
                {
                    if (nbt.nid.nidType == EnidType.NORMAL_MESSAGE)
                    {
                        RebuildMessageObject(nbt);
                    }
                    else
                    {
                        if (NDB.IsPC(nbt.nid))
                        {
                            RebuildPC(nbt);
                        }
                        else
                        {
                            if (NDB.IsTC(nbt.nid))
                            {
                                RebuildTC(nbt);
                            }
                            else
                            {
                                outFile.CopyNDB(nbt);
                                if (nbt.nid.nidType == EnidType.CONTENTS_TABLE_INDEX)
                                {
                                    unsafe
                                    {
                                        outFile.header.rgnid[3] = outFile.header.rgnid[19];  // strange error in scanpst
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        public static Folder CheckFoldersToExport(string folderToExport)
        {
            GetFolderList();
            Folder expFolder = SelectFoldersToExport(folderToExport);
            NBTENTRY nbtFolder = srcFile.NBTs[expFolder.nbtIndex];
            nbtFolder.nidParent = RQS.NID_TOP_PERSONAL_FOLDER.dwValue;
            srcFile.NBTs[expFolder.nbtIndex] = nbtFolder;
            expFolder.properties = LTP.ReadPCs(srcFile.stream, nbtFolder);
            return expFolder;
        }
        private static Folder SelectFoldersToExport(string topFolder)
        {
            int fIx = folders.FindIndex(f => f.name == topFolder);
            if (fIx < 0)
            {  // need to confirm whether deleled items folder is always present
                fIx = folders.FindIndex(f => f.path == topFolder);
                if (fIx < 0)
                {
                    throw new Exception($"Selected folder: {topFolder} not found in the OST file");
                }
            }
            MarkSubfoldersToExport(folders[fIx]); ;
            return folders[fIx];
        }
        private static void MarkSubfoldersToExport(Folder folder)
        {
            folder.toBeExported = true;
            if (folder.name == "IPM_COMMON_VIEWS") return;  // scanpst remo
            foreach (Folder f in folders)
            {
                if (f.parent.nid.dwValue == folder.nid.dwValue)
                {
                    MarkSubfoldersToExport(f);
                }
            }
        }
        public static void GetFolderList()
        {
            Program.mainForm.statusMSG($"Scanning ost folder list....");
            folders = new List<Folder>();
            for (int i = 0; i < srcFile.NBTs.Count; i++)
            {
                NBTENTRY nbt = srcFile.NBTs[i];
                if (nbt.nid.dwValue == (uint)EnidSpecial.NID_ROOT_FOLDER)
                {
                    Folder folder = new Folder()
                    {
                        name = "",
                        path = @"\",
                        level = 0,
                        nid = nbt.nid,
                        toBeExported = false,
                        nbtIndex = i
                    };
                    folder.parent = folder;
                    folders.Add(folder);
                    RootFolder = folder;
                }
                else
                {
                    if (nbt.nid.nidType == EnidType.NORMAL_FOLDER | nbt.nid.nidType == EnidType.SEARCH_FOLDER)
                    {
                        Folder folder = new Folder()
                        {
                            nbtIndex = i,
                            name = ReadFolderName(nbt),
                            parent = parentFolder(nbt),
                            nid = nbt.nid,
                            toBeExported = false,
                        };
                        if (folder.name == "IPM_SUBTREE")
                        {
                            folder.path = "";
                        }
                        else
                        {
                            folder.path = folder.parent.path + @"\" + folder.name;
                        }
                        folders.Add(folder);
                    }
                }
            }
        }
        private static Folder parentFolder(NBTENTRY nbt)
        {
            int folderIx = folders.FindIndex(f => f.nid.dwValue == nbt.nidParent);
            if (folderIx < 0) throw new Exception($"Parent folder NID {nbt.nidParent} not found");
            return folders[folderIx];
        }
        private static string ReadFolderName(NBTENTRY nbt)
        {
            string folderName = "";
            List<Property> props = ReadAndValidateOSTpcs(nbt);
            int nameIx = props.FindIndex(p => (p.PCBTH.wPropId == EpropertyId.PidTagDisplayName & p.PCBTH.wPropType == EpropertyType.PtypString));
            if (nameIx >= 0)
            {
                byte[] name = props[nameIx].data;
                folderName = Encoding.Unicode.GetString(name);
            }
            return folderName;
        }
        private static bool ToBeExported(NBTENTRY nbt)
        {
            if (nbt.nid.dwValue == 34666) return false;

            if (nbt.bidData == 0) return false;
            if (outFile.NBTs.FindIndex(n => n.nid.dwValue == nbt.nid.dwValue) >= 0)
            {
                return false;
            }
            if (IsFolderData(nbt))
            {
                NID fNid = nbt.nid.ChangeType(EnidType.NORMAL_FOLDER);
                return IsFolderToBeExported(fNid.dwValue);
            }
            else
            {
                if (nbt.nid.nidType == EnidType.NORMAL_MESSAGE || nbt.nid.nidType == EnidType.ASSOC_MESSAGE)
                {
                    return IsFolderToBeExported(nbt.nidParent);
                }
                else
                {
                    return IsSpecialNidToExport(nbt);
                }
            }
        }
        private static bool IsSpecialNidToExport(NBTENTRY nbt)
        {
            switch (nbt.nid.dwValue)
            {
                case 0x0061:    // NID_NAME_TO_ID_MAP
                case 0x06B6:    // unknown type but required
                    return true;
                default: return false;
            }

        }
        private static bool IsFolderToBeExported(UInt32 nid)
        {
            int fIx = folders.FindIndex(f => f.nid.dwValue == nid);
            if (fIx >= 0)
            {
                return folders[fIx].toBeExported;
            }
            return false;
        }
        private static bool IsFolderData(NBTENTRY nbt)
        {
            switch (nbt.nid.nidType)
            {
                case EnidType.NORMAL_FOLDER:
                case EnidType.HIERARCHY_TABLE:
                case EnidType.CONTENTS_TABLE:
                case EnidType.ASSOC_CONTENTS_TABLE:
                case EnidType.CONTENTS_TABLE_INDEX:
                    return true;
                default: return false;
            }
        }
        public static void exportNBTnodes()

        {   // NBT's should be the same
            outFile.NBTs = (List<NBTENTRY>)outFile.NBTs.OrderBy(h => h.nid.dwValue).ToList();
            NTree nodeTree = new NTree(outFile.NBTs);
            outFile.header.root.BREFNBT = nodeTree.ExportNodes(outFile);
        }
        public static void exportBBTnodes()
        {
            outFile.BBTs = (List<BBTENTRY>)outFile.BBTs.OrderBy(h => h.BREF.bid).ToList();
            NTree nodeTree = new NTree(outFile.BBTs);
            outFile.header.root.BREFBBT = nodeTree.ExportNodes(outFile);
        }
        public static void updateNidHighWaterMarks()
        {
            for (int i = 0; i < outFile.NBTs.Count; i++)
            {
                UpdateNIDheaderIndex(outFile.NBTs[i].nid.dwValue);
            }
        }
        public static unsafe void UpdateNIDheaderIndex(UInt32 nid)
        {
            int nidType = (int)nid % 32;
            uint nidIx = nid >> 5;
            if (nidIx > outFile.header.rgnid[nidType]) { outFile.header.rgnid[nidType] = nidIx; }
        }
        public static void UpdateSubnodeNIDheaderIndex(UInt32 nid)
        {   // PSTSCAN seems to check the highmarks of specific subnode nid type
            // this is not documented and outlook don't have a problem with it
            int nidType = (int)nid % 32;
            switch (nidType)
            {
                case 1:
                case 3:
                case 4:
                case 5:
                case 8:
                case 9:
                case 13:
                // case 18:
                case 19:
                case 20:
                case 21:
                case 22:

                case 24:
                case 27:
                case 29:
                case 31:
                    UpdateNIDheaderIndex(nid);
                    break;
            }
        }
        public static string SourceFileDetails()
        {
            return fileSummary(srcFile);
        }
        public static string OutputFileDetails()
        {
            return fileSummary(outFile);
        }
        public static UInt32 NextUniqueID()
        {   // outlook seems to use this value for 
            // PidTagLtpRowVer property on table context on outlook
            // this is not documented
            UInt32 nextValue = outFile.header.dwUnique;                     //// TO BE CONFIRMED
            outFile.header.dwUnique++;
            return nextValue;
        }
        private static string fileSummary(XstFile xstFile)
        {
            string fd = "";
            fd = fd + $"Filename        : {xstFile.filename} " + Environment.NewLine;
            fd = fd + $"File Type       : {xstFile.type.ToString()}" + Environment.NewLine;
            fd = fd + $"File Size       : {xstFile.header.root.ibFileEof / 1024}kb" + Environment.NewLine;
            fd = fd + $"Free area       : {xstFile.header.root.cbAMapFree / 1024}kb" + Environment.NewLine;
            fd = fd + $"bCryptMethod    : {xstFile.header.bCryptMethod}" + Environment.NewLine;
            fd = fd + $"Nr of NBTs      : {xstFile.NBTs.Count}" + Environment.NewLine; ;
            if (xstFile.type == FileType.OST)
            {
                fd = fd + $"Nr of BBTs      : {xstFile.BBTs.Count}"; ;
            }
            else
            {
                fd = fd + $"Nr of BBTs      : {xstFile.BBTs.Count}"; ;
            }
            return fd;
        }
        private static string fileSummary(PstFile xstFile)
        {
            string fd = "";
            fd = fd + $"Filename        : {xstFile.filename} " + Environment.NewLine;
            fd = fd + $"File Type       : PST" + Environment.NewLine;
            fd = fd + $"File Size       : {xstFile.header.root.ibFileEof / 1024}kb" + Environment.NewLine;
            fd = fd + $"Free area       : {xstFile.header.root.cbAMapFree / 1024}kb" + Environment.NewLine;
            fd = fd + $"bCryptMethod    : {xstFile.header.bCryptMethod}" + Environment.NewLine;
            fd = fd + $"Nr of NBTs      : {xstFile.NBTs.Count}" + Environment.NewLine; ;
            fd = fd + $"Nr of BBTs      : {xstFile.BBTs.Count}";
            return fd;
        }
        public static void CloseSourceFile()
        {
            srcFile?.Close();
        }
        public static void CloseOutputFile()
        {
            outFile?.FinalizePSTdata();
            outFile?.Close();
        }
        #region file stream handlers
        // methods for
        // - reading/writing OST/PST byter arrays to/from c# class types
        // - converting (Marshal) byte arrays
        public static uint PageCRC<T>(T pageData)
        {   // MS_PST section 2.2.2.7.1	PAGETRAILER
            int dataBytes = MS.AMapPageEntryBytes;
            byte[] array = new byte[dataBytes];
            ExtractArrayFromType<T>(ref array, pageData, dataBytes);
            return PST.ComputeCRC(array, dataBytes);
        }

        // 
        // extracts given data type from a byte array
        public static unsafe T ExtractTypeFromArray<T>(byte[] array, int offset = 0)
        {
            int size = Marshal.SizeOf(typeof(T));
            if (offset < 0 || offset + size > array.Length)
                throw new Exception($"ExtractTypeFromArray Error: data type {typeof(T).Name} at at offset {offset} out of array range length {array.Length}");
            GCHandle handle = GCHandle.Alloc(array, GCHandleType.Pinned);
            T dataType = (T)Marshal.PtrToStructure(handle.AddrOfPinnedObject() + offset, typeof(T));
            handle.Free();
            return dataType;
        }
        public static void ExtractArrayFromType<T>(ref byte[] array, T typeData, int length, int tPos = 0, int bPos = 0)
        {
            int size = Marshal.SizeOf(typeData);
            byte[] dataBytes = new byte[size];
            IntPtr ptr = IntPtr.Zero;
            try
            {
                ptr = Marshal.AllocHGlobal(size);
                Marshal.StructureToPtr(typeData, ptr, true);
                Marshal.Copy(ptr, dataBytes, 0, size);
                Array.Copy(dataBytes, tPos, array, bPos, length);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }
        public static T[] ArrayFromArray<T>(byte[] array, int tPos, int tCount)
        {
            T[] nArray = new T[tCount];
            int size = Marshal.SizeOf(typeof(T));
            if (tPos < 0 || tPos + tCount * size > array.Length)
                throw new Exception($"Extract Array from another Array error: {array.Length} at offset {tPos}");
            GCHandle handle = GCHandle.Alloc(array, GCHandleType.Pinned);
            for (int i = 0; i < tCount; i++)
            {
                nArray[i] = (T)Marshal.PtrToStructure(handle.AddrOfPinnedObject() + tPos + i * size, typeof(T));
            }
            handle.Free();
            return nArray;
        }
        public static unsafe T[] MapArray<T>(byte[] buffer, int offset, int count)
        {
            T[] temp = new T[count];
            int size = Marshal.SizeOf(typeof(T));
            if (offset < 0 || offset + count * size > buffer.Length)
                throw new Exception($"Out of bounds error attempting to map array of {count} {typeof(T).Name}s from byte array of length {buffer.Length} at offset {offset}");
            GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            for (int i = 0; i < count; i++)
            {
                temp[i] = (T)Marshal.PtrToStructure(handle.AddrOfPinnedObject() + offset + i * size, typeof(T));
            }
            handle.Free();
            return temp;
        }
        public static unsafe void TypeToArray<T>(ref byte[] buffer, T typeData, int offset = 0)  /// Nivaldo
        {
            int size = Marshal.SizeOf(typeof(T));
            byte[] dataBytes = new byte[size];
            IntPtr ptr = IntPtr.Zero;
            try
            {
                ptr = Marshal.AllocHGlobal(size);
                Marshal.StructureToPtr(typeData, ptr, true);
                Marshal.Copy(ptr, dataBytes, 0, size);
                Array.Copy(dataBytes, 0, buffer, offset, size);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }
        public unsafe static T TypeFromBuffer<T>(byte* buffer, int buflen, int offset)
        {
            int size = Marshal.SizeOf(typeof(T));
            if (offset < 0 || offset + size > buflen)
                throw new Exception($"ExtractTypeFromArray Error: data type {typeof(T).Name} at at offset {offset} out of array range length {buflen}");
            T dataType = (T)Marshal.PtrToStructure(new IntPtr(buffer + offset), typeof(T));
            return dataType;
        }
        public unsafe static void TypeToBuffer<T>(byte* buffer, int buflen, T typeData, int offset = 0)  /// Nivaldo
        {
            int size = Marshal.SizeOf(typeof(T));
            { // it just copy the nr of bytes in buffer 
                size = buflen - offset;
            }
            Marshal.StructureToPtr(typeData, new IntPtr(buffer + offset), true);
        }
        public static T ReadType<T>(FileStream fs)
        {
            byte[] array = new byte[Marshal.SizeOf(typeof(T))];
            fs.Read(array, 0, Marshal.SizeOf(typeof(T)));
            return ExtractTypeFromArray<T>(array);
        }
        public static T ReadType<T>(FileStream fs, UInt64 offset)
        {
            fs.Position = (long)offset;
            return ReadType<T>(fs);
        }
        public static void WriteType<T>(FileStream fs, T typeData, UInt64 offset = 0)
        {
            int size = Marshal.SizeOf(typeof(T));
            byte[] buffer = new byte[size];
            IntPtr ptr = IntPtr.Zero;
            try
            {
                ptr = Marshal.AllocHGlobal(size);
                Marshal.StructureToPtr(typeData, ptr, true);
                Marshal.Copy(ptr, buffer, 0, size);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
            if (offset > 0) { fs.Position = (long)offset; }
            fs.Write(buffer, 0, size);
        }
        #endregion
    }
}