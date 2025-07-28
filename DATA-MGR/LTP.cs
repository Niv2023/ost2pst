// LTP layer for OST/PST file handling

using System.Runtime.InteropServices;

namespace ost2pst
{
    public class HeapItem
    {
        private HID _hid;
        private byte[] _data;
        public int Size { get { if (_data is null) return 0; return _data.Length; } }
        public HID Hid { get { return _hid; } set { _hid = value; } }
        public void HID(int hidIndex, int hidBlokIndex)
        {
            UInt16 hIx = (UInt16)(hidIndex << 5);
            UInt16 hBlok = (UInt16)hidBlokIndex;
            _hid = new HID(hIx, hBlok);
        }
        public byte[] Data { get { return _data; } }
        public HeapItem()
        {
            _data = null;
            _hid = new HID(0, 0);
        }
        public HeapItem(HID hid)
        {
            _hid = hid;
            _data = null;
        }
        public HeapItem(HID hid, byte[] data)
        {
            _hid = hid;
            _data = data;
        }
        public HeapItem(byte[] data)
        {
            _hid = new HID(0, 0);
            _data = data;
        }
    }

    public static class LTP
    {

        // Generation replication item ids
        // not documented.... checked what SCANPST generates and some clues from chatgpt
        public static byte[] ReplItemId = new byte[16];
        public static UInt64 ReplChangeNumber = 0x100;
        public static byte[] ReplVersionHistory = new byte[24] { 1, 0, 0, 0, 243, 136, 137, 115, 10, 69, 209, 70, 151, 195, 23, 245, 7, 39, 50, 56, 1, 0, 0, 0 };

       private static void UpdateReplTags()
        {
            ReplItemId = Guid.NewGuid().ToByteArray();
            ReplChangeNumber += 8;
        }
        public static List<Property> ReplTagProps ()
        {
            List<Property> props = new List<Property>();
            UpdateReplTags();
            Property replItemId = new Property(EpropertyId.PidTagReplItemid, EpropertyType.PtypBinary, ReplItemId);
            props.Add(replItemId);
            Property replChangeNumber = new Property(EpropertyId.PidTagReplChangenum, EpropertyType.PtypInteger64, ReplChangeNumber);
            props.Add(replChangeNumber);
            Property replVersionHistory = new Property(EpropertyId.PidTagReplVersionHistory, EpropertyType.PtypBinary, ReplVersionHistory);
            props.Add(replVersionHistory);
            return props;
        } 

        // A heap-on-node data block

        public static int largestHID;
        private class HNDataBlock
        {
            public int Index;
            public byte[] Buffer;
            public UInt16[] rgibAlloc;

            // In first block only
            public EbType bClientSig;
            public HID hidUserRoot;  // HID that points to the User Root record
        }
        // Used when reading table data to normalise handling of in-line and sub node data storage
        private class RowDataBlock
        {
            public byte[] Buffer;
            public int Offset;
            public int Length;
        }
        #region write functions
        public static List<byte[]> GetBTHhnDatablocks(BTH bth, EbType type)
        {
            List<byte[]> dbs = new List<byte[]>();
            // first hn block is the BTH header
            for (int i = 0; i <= bth.hidBlockIx; i++)
            {
                byte[] db = HNdatablock(bth, i, type);
                dbs.Add(db);
            }
            return dbs;
        }
        private static byte[] HNdatablock(BTH bth, int blockIx, EbType type)
        {
            List<HeapItem> hItems = bth.heapItems.FindAll(h => h.Hid.GetBlockIndex(false) == blockIx);
            byte[] hnBlock = NewHNdatablock(blockIx, bth.hidBlockIx, type);
            List<UInt16> rgibAlloc = new List<UInt16>() { (UInt16)hnBlock.Length };
            foreach (HeapItem h in hItems)
            {   // the list should be sorted on hids
                int iPos = hnBlock.Length;
                Array.Resize(ref hnBlock, hnBlock.Length + h.Size);
                Array.Copy(h.Data, 0, hnBlock, iPos, h.Size);
                rgibAlloc.Add((UInt16)hnBlock.Length);
            }
            bool lastBlock = bth.hidBlockIx == blockIx;
            CloseHNdatablock(ref hnBlock, rgibAlloc, lastBlock);
            return hnBlock;
        }

        private static byte[] NewHNdatablock(int blockIndex, int lastIndex, EbType bClientSig = 0)  // EbType only for blockIx=0
        {  // just returns the HN datablock plain header
            byte[] hnBlock = new byte[0];
            int nrBlocks = lastIndex - blockIndex + 1;
            if (blockIndex == 0)
            { // HNHDR
                hnBlock = new byte[12];
                hnBlock[2] = 0xEC;
                hnBlock[3] = (byte)bClientSig;
                HID hidUserRoot = new HID(EnidType.HID, 1, 0);
                if (bClientSig == EbType.bTypeTC)
                {
                    hidUserRoot = new HID(EnidType.HID, 2, 0);
                }
                FM.TypeToArray<UInt32>(ref hnBlock, hidUserRoot.hidValue, 4);
                if (nrBlocks > 8) { nrBlocks = 8; } // max 8 blocks in the first rgbFillLevel
                rgbFillLevel(hnBlock, 8, nrBlocks);
                return hnBlock;
            }
            if ((blockIndex - 8) % 128 == 0)
            { // HNBITMAPHDR
                hnBlock = new byte[66];
                if (nrBlocks > 128) { nrBlocks = 128; }
                rgbFillLevel(hnBlock, 2, nrBlocks);
                return hnBlock;
            }
            // HNPAGEDR
            hnBlock = new byte[2];
            return hnBlock;
        }
        private static void rgbFillLevel(byte[] hnBlock, int rgbIx, int nrBlocks)
        {   // all created HN blocks are fully allocated so FFFF....
            int nEntries = nrBlocks / 2;
            for (int i = 1; i <= nEntries; i++)
            {
                hnBlock[rgbIx + i - 1] = 0xff;
            }
            if (nrBlocks % 2 != 0)
            {
                hnBlock[rgbIx + nEntries] = 0x0f;
            }
        }
        private static void CloseHNdatablock(ref byte[] hnBlock, List<UInt16> rgibAlloc, bool lastBlock)
        {
            int hnSize = hnBlock.Length;
            if (hnSize % 2 != 0) { hnSize++; }  // align 2 bytes for the HNPAGEMAP
            UInt16 ibHnpm = (UInt16)hnSize;
            //hnSize += rgibAlloc.Count * 2;  // cAlloc + cFree + rgibAlloc bytes
            hnSize += rgibAlloc.Count * 2 + 4;  // cAlloc + cFree + rgibAlloc bytes
            if (lastBlock)
            {
                Array.Resize(ref hnBlock, hnSize);
            }
            else
            {  // blocks in a xtree should have a fixed max size
                Array.Resize(ref hnBlock, (int)MS.MaxDatablockSize);
            }
            FM.TypeToArray<UInt16>(ref hnBlock, ibHnpm);
            UInt16 cAlloc = (UInt16)(rgibAlloc.Count - 1);
            FM.TypeToArray<UInt16>(ref hnBlock, cAlloc, ibHnpm);
            FM.TypeToArray<UInt16>(ref hnBlock, 0, ibHnpm + 2);   // cFree = 0
            for (int i = 0; i < rgibAlloc.Count; i++)
            {
                FM.TypeToArray<UInt16>(ref hnBlock, rgibAlloc[i], ibHnpm + 4 + i * 2);
            }
        }
        #endregion
        #region read functions
        public static List<Property> ReadPCs(FileStream fs, NBTENTRY nbt)
        {
            List<Property> properties = new List<Property>();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(nbt.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (nbt.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(nbt.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypePC)
                throw new Exception($"NID entry {nbt.nid.dwValue} is not a PC");
            properties = GetPropertiesFromHeap(heap, nidSubnodes);
            if (nbt.nid.dwValue == 8412004)
            {
                EpropertyId p = properties[9].id;
                byte[] data = properties[9].data;
            }
            return properties;
        }
        public static List<Property> ReadSubnodePC(FileStream fs, SLENTRY sn)
        {
            List<Property> properties = new List<Property>();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(sn.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (sn.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(sn.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypePC)
                throw new Exception($"NID entry {sn.nid} is not a PC");
            properties = GetPropertiesFromHeap(heap, nidSubnodes);
            return properties;
        }
        public static TableContext ReadSubnodeTC(FileStream fs, SLENTRY sn)
        {
            TableContext TC = new TableContext();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(sn.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (sn.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(sn.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypeTC)
                throw new Exception($"NID entry {sn.nid} is not a TC");
            TC.tcINFO = MapType<TCINFO>(heap, h.hidUserRoot);
            // Read the column descriptions
            TC.tcCols = MapArray<TCOLDESC>(heap, h.hidUserRoot, TC.tcINFO.cCols, Marshal.SizeOf(typeof(TCINFO))).ToList();
            TC.tcRowIndexes = GetTcBTHIndex(heap, TC.tcINFO.hidRowIndex);
            TC.tcRowMatrix = new List<RowData>();
            TC.tcRowMatrix = GetRowData(TC, heap, nidSubnodes);
            return TC;
        }
        public static TableContext ReadSubnodeTCandRefreshIndex(FileStream fs, SLENTRY sn)
        {
            TableContext TC = new TableContext();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(sn.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (sn.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(sn.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypeTC)
                throw new Exception($"NID entry {sn.nid} is not a TC");
            TC.tcINFO = MapType<TCINFO>(heap, h.hidUserRoot);
            // Read the column descriptions
            TC.tcCols = MapArray<TCOLDESC>(heap, h.hidUserRoot, TC.tcINFO.cCols, Marshal.SizeOf(typeof(TCINFO))).ToList();
            TC.tcRowIndexes = GetTcBTHIndex(heap, TC.tcINFO.hidRowIndex);
            TC.tcRowMatrix = new List<RowData>();
            TC.tcRowMatrix = GetRowData(TC, heap, nidSubnodes);
            for (int i = 0; i < TC.tcRowIndexes.Count; i++)
            {
                TCROWID rId = new TCROWID() { dwRowID = TC.tcRowIndexes[i].dwRowID, dwRowIndex = (UInt32)i };
                TC.tcRowIndexes[i] = rId;
            }
            return TC;
        }
        public static TableContext ReadTCs(FileStream fs, NBTENTRY nbt)
        {
             return ReadTCs_new_rowdata(fs, nbt);
        }
        public static TableContext ReadTCs_and_rowdata(FileStream fs, NBTENTRY nbt)
        {
            TableContext TC = new TableContext();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(nbt.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (nbt.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(nbt.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypeTC)
                throw new Exception($"NID entry {nbt.nid.dwValue} is not a TC");
            TC.tcINFO = MapType<TCINFO>(heap, h.hidUserRoot);
            // Read the column descriptions
            TC.tcCols = MapArray<TCOLDESC>(heap, h.hidUserRoot, TC.tcINFO.cCols, Marshal.SizeOf(typeof(TCINFO))).ToList();
            TC.tcRowIndexes = GetTcBTHIndex(heap, TC.tcINFO.hidRowIndex);
            TC.tcRowMatrix = GetRowData(TC, heap, nidSubnodes);
            // row matrix should be ordered by rowid
            // reset the indexes since the row matrix is built as it is now
            for (int i = 0; i < TC.tcRowIndexes.Count; i++)
            {
                TCROWID rId = new TCROWID() { dwRowID = TC.tcRowIndexes[i].dwRowID, dwRowIndex = (UInt32)i };
                TC.tcRowIndexes[i] = rId;
            }
            return TC;
        }
        public static TableContext ReadTCs_new_rowdata(FileStream fs, NBTENTRY nbt)
        {
            TableContext TC = new TableContext();
            List<byte[]> nidDatablocks = FM.srcFile.ReadDatablocks(nbt.bidData);
            List<SLENTRY> nidSubnodes = new();
            if (nbt.bidSub > 0) nidSubnodes = FM.srcFile.GetSLentries(nbt.bidSub);
            List<HNDataBlock> heap = ReadHeapOnNode(fs, nidDatablocks);
            var h = heap.First();
            if (h.bClientSig != EbType.bTypeTC)
                throw new Exception($"NID entry {nbt.nid.dwValue} is not a TC");
            TC.tcINFO = MapType<TCINFO>(heap, h.hidUserRoot);
            // Read the column descriptions
            TC.tcCols = MapArray<TCOLDESC>(heap, h.hidUserRoot, TC.tcINFO.cCols, Marshal.SizeOf(typeof(TCINFO))).ToList();
            TC.tcRowIndexes = GetTcBTHIndex(heap, TC.tcINFO.hidRowIndex);
            // rebuild the TC table since the OST PCs seems not be proper reflected in the row data
            TC.tcRowMatrix = new List<RowData>();
            for (int i = 0; i < TC.tcRowIndexes.Count; i++)
            {
                TCROWID rId = new TCROWID() { dwRowID = TC.tcRowIndexes[i].dwRowID, dwRowIndex = (UInt32)i };
                TC.tcRowIndexes[i] = rId;
                int pcNbt = FM.srcFile.NBTs.FindIndex(n => n.nid.dwValue == rId.dwRowID);
                if (pcNbt < 0) throw new Exception($"TC (nid={nbt.nid.dwValue}) refers to an unknown id {rId.dwRowID}");
                List<Property> props = ReadPCs(fs, FM.srcFile.NBTs[pcNbt]);
                if (FM.srcFile.NBTs[pcNbt].nid.nidType == EnidType.NORMAL_FOLDER) RQS.ValidateFolderPC(props);
                else if (FM.srcFile.NBTs[pcNbt].nid.nidType == EnidType.NORMAL_MESSAGE || FM.srcFile.NBTs[pcNbt].nid.nidType == EnidType.ASSOC_MESSAGE) RQS.ValidateMessagePC(props);
                props.AddRange(LTP.ReplTagProps());
                props = props.OrderBy(h => h.PCBTH.wPropId).ToList();
                TC.AddRow(rId, props);
            }
            return TC;
        }
        private static List<RowData> GetRowData(TableContext TC, List<HNDataBlock> blocks, List<SLENTRY> subnodes)
        {
            List<RowData> rowsData = new List<RowData>();
            if (TC.tcINFO.hnidRows.HasValue)
            {
                List<RowDataBlock> db = new List<RowDataBlock>();
                if (TC.tcINFO.hnidRows.IsHID)
                {
                    // Raw data within a heap item
                    var buf = GetBytesForHNID(blocks, subnodes, TC.tcINFO.hnidRows);
                    db.Add(
                        new RowDataBlock
                        {
                            Buffer = buf,
                            Offset = 0,
                            Length = buf.Length,
                        });
                }
                else  // should be a NID
                {
                    db = ReadSubNodeRowDataBlocks(subnodes, TC.tcINFO.hnidRows.NID);
                }
                for (int i = 0; i < TC.tcRowIndexes.Count; i++)
                {
                    List<Property> props = ReadColumnValues(TC, blocks, subnodes, db, TC.tcRowIndexes[i].dwRowIndex);
                    RowData rowData = new RowData(TC.tcINFO, TC.tcRowIndexes[i].dwRowID, props);
                    rowsData.Add(rowData);
                }
            }
            return rowsData;
        }
        private static List<Property> ReadColumnValues(TableContext TC, List<HNDataBlock> blocks, List<SLENTRY> subnodes, List<RowDataBlock> dbs, UInt32 rowIndex)
        {
            List<Property> properties = new List<Property>();
            int rgCEBSize = (int)Math.Ceiling((decimal)TC.tcINFO.cCols / 8);
            int rowsPerBlock = ((int)MS.MaxBlockSize - Marshal.SizeOf(typeof(BLOCKTRAILER))) / TC.tcINFO.rgibTCI_bm;
            if (FM.srcFile.type == FileType.OST)
            {
                rowsPerBlock = (MS.MaxOSTBlockSize - Marshal.SizeOf(typeof(OST_BLOCKTRAILER))) / TC.tcINFO.rgibTCI_bm;
            }
            int bn = (int)(rowIndex / rowsPerBlock);
            if (bn >= dbs.Count) throw new Exception("Data block number out of bounds");
            var db = dbs[bn];
            long rowOffset = db.Offset + (rowIndex % rowsPerBlock) * TC.tcINFO.rgibTCI_bm;
            // Read the column existence data
            var rgCEB = FM.MapArray<Byte>(db.Buffer, (int)(rowOffset + TC.tcINFO.rgibTCI_1b), rgCEBSize);
            foreach (var col in TC.tcCols)
            {
                // Check if the column exists
                if ((rgCEB[col.iBit / 8] & (0x01 << (7 - (col.iBit % 8)))) == 0)
                    continue;
                Property prop = new Property(col.wPropId, col.wPropType);
                try
                {
                    prop = ReadTableCell(blocks, subnodes, db, rowOffset, col);
                    if (prop.type == EpropertyType.PtypString & prop.dataSize == 0)
                    {  
                        // SCANPST 'sometimes' gives error for empty string cells... outlook seems to accept
                    }
                    else
                    {
                        properties.Add(prop);
                    }
                }
                catch
                {
                    Program.mainForm.statusMSG($"Unable to get the value for this table column value: prop id={col.wPropId}, prop type={col.wPropType}");
                }
            }
            return properties;
        }
        private static List<RowDataBlock> ReadSubNodeRowDataBlocks(List<SLENTRY> subnodes, NID nid)
        {
            var blocks = new List<RowDataBlock>();
            int index = subnodes.FindIndex(h => h.nid == nid.dwValue);
            if (index < 0) throw new Exception("Sub node NID not found");
            foreach (var buf in FM.srcFile.ReadDatablocks(subnodes[index].bidData))
            {
                blocks.Add(new RowDataBlock
                {
                    Buffer = buf,
                    Offset = 0,
                    Length = buf.Length,
                });
            }
            return blocks;
        }

        private static List<Property> GetPropertiesFromHeap(List<HNDataBlock> heap, List<SLENTRY> subnodes)
        {
            List<Property> props = new List<Property>();
            var h = heap.First();
            BTHHEADER bh = MapType<BTHHEADER>(heap, h.hidUserRoot);
            foreach (PCBTH pcbth in ReadBTHIndexHelper<PCBTH>(heap, bh.hidRoot, bh.bIdxLevels))
            {
                byte[] data = new byte[0];
                switch (pcbth.wPropType)
                {
                    case EpropertyType.PtypInteger16:
                        data = new byte[2];
                        FM.TypeToArray<UInt16>(ref data, (UInt16)pcbth.dwValueHnid.dwValue);
                        break;
                    case EpropertyType.PtypInteger32:
                        data = new byte[4];
                        FM.TypeToArray<UInt32>(ref data, (UInt32)pcbth.dwValueHnid.dwValue);
                        break;
                    case EpropertyType.PtypBoolean:
                        data = new byte[1] { (byte)pcbth.dwValueHnid.dwValue };
                        break;
                    default:
                        if (!pcbth.dwValueHnid.HasValue) continue;
                        {
                            if (IsDataOnHNID(pcbth))
                            {
                                data = GetDataFormHNID(pcbth);
                            }
                            else
                            {
                                if (pcbth.dwValueHnid.IsHID)
                                {   // data on the heap
                                    data = GetBytesForHNID(heap, subnodes, pcbth.dwValueHnid);
                                }
                                else
                                {   // data in a subnode
                                    NID sNid = new NID(pcbth.dwValueHnid.dwValue);
                                    SLENTRY sn = subnodes.Find(h => h.nid == sNid.dwValue);
                                    data = FM.srcFile.ReadFullDatablock(sn.bidData);
                                }
                            }
                        }
                        break;
                }
                Property prop = new Property(pcbth, data);
                props.Add(prop);
            }
            return props;
        }
        private static byte[] GetDataFormHNID(PCBTH pcbth)
        {  // for fixed sized PCBTH
            byte[] data = new byte[4];
            if (pcbth.wPropType == EpropertyType.PtypBoolean)
            {
                data = new byte[1];
                data[0] = (byte)pcbth.dwValueHnid.dwValue; 
            }
            else if (pcbth.wPropType == EpropertyType.PtypInteger16)
            {
                data = new byte[2];
                FM.TypeToArray<UInt16>(ref data, (UInt16)pcbth.dwValueHnid.dwValue);
            }
            else
            {
                FM.TypeToArray<UInt32>(ref data, pcbth.dwValueHnid.dwValue);
            }
            return data;
        }
        private static bool IsDataOnHNID(PCBTH pcbth)
        {   // true for fixed types <= 4 bytes
            bool result = false;
            switch (pcbth.wPropType)
            {
                // 1 byte types
                case EpropertyType.PtypBoolean:
                case EpropertyType.PtypInteger16:
                case EpropertyType.PtypInteger32:
                    result = true;
                    break;
                    // 8 byte types
            }
            return result;
        }
        private static List<TCROWID> GetTcBTHIndex(List<HNDataBlock> heap, HID hid)
        {
            List<TCROWID> rowIDs = new List<TCROWID>();
            BTHHEADER bh = MapType<BTHHEADER>(heap, hid);
            foreach (TCROWID rowID in ReadBTHIndexHelper<TCROWID>(heap, bh.hidRoot, bh.bIdxLevels))
            {
                rowIDs.Add(rowID);
            }
            return rowIDs;
        }

        private static IEnumerable<T> ReadBTHIndexHelper<T>(List<HNDataBlock> blocks, HID hid, int level)
        {
            if (level == 0)
            {
                int recCount = HidSize(blocks, hid) / Marshal.SizeOf(typeof(T));
                if (hid.GetIndex(FM.srcFile.type == FileType.OST) != 0)
                {
                    // The T record also forms the key of the BTH entry
                    foreach (var row in MapArray<T>(blocks, hid, recCount))
                        yield return row;
                }
            }
            else
            {
                int recCount = HidSize(blocks, hid) / Marshal.SizeOf(typeof(IntermediateBTH4));
                var inters = MapArray<IntermediateBTH4>(blocks, hid, recCount);

                foreach (var inter in inters)
                {
                    foreach (var row in ReadBTHIndexHelper<T>(blocks, inter.hidNextLevel, level - 1))
                        yield return row;
                }
            }
            yield break; // No more entries
        }
        private static int HidSize(List<HNDataBlock> blocks, HID hid)
        {

            var index = hid.GetIndex(FM.srcFile.type == FileType.OST);
            if (index == 0) // Check for empty
                return 0;
            var b = blocks[hid.GetBlockIndex(FM.srcFile.type == FileType.OST)];
            int hidSize = b.rgibAlloc[index] - b.rgibAlloc[index - 1];
            if (hidSize > largestHID)
            {
                largestHID = hidSize;
            }
            return hidSize;
        }
        private static List<HNDataBlock> ReadHeapOnNode(FileStream fs, List<byte[]> datablocks)
        {
            var blocks = new List<HNDataBlock>();
            int index = 0;
            foreach (var buf in datablocks)
            {
                // First block contains a HNHDR
                byte[] block = buf;
                if (index == 0)
                {
                    var h = FM.ExtractTypeFromArray<HNHDR>(buf, 0);
                    var pm = FM.ExtractTypeFromArray<HNPAGEMAP>(buf, h.ibHnpm);
                    var b = new HNDataBlock
                    {
                        Index = index,
                        Buffer = buf,
                        bClientSig = h.bClientSig,
                        hidUserRoot = h.hidUserRoot,
                        rgibAlloc = FM.ArrayFromArray<UInt16>(block, (ushort)(h.ibHnpm + Marshal.SizeOf(pm)), pm.cAlloc + 1),  //+1 to get the dummy entry that gives us the size of the last one
                    };
                    blocks.Add(b);
                }
                // Blocks 8, 136, then every 128th contains a HNBITMAPHDR
                else if (index == 8 || (index >= 136 && (index - 8) % 128 == 0))
                {
                    var h = FM.ExtractTypeFromArray<HNBITMAPHDR>(buf, 0);
                    var pm = FM.ExtractTypeFromArray<HNPAGEMAP>(buf, h.ibHnpm);
                    var b = new HNDataBlock
                    {
                        Index = index,
                        Buffer = buf,
                        rgibAlloc = FM.ArrayFromArray<UInt16>(buf, h.ibHnpm + Marshal.SizeOf(pm), pm.cAlloc + 1),  //+1 to get the dummy entry that gives us the size of the last one
                    };
                    blocks.Add(b);
                }
                // All other blocks contain a HNPAGEHDR
                else
                {
                    var h = FM.ExtractTypeFromArray<HNPAGEHDR>(buf, 0);
                    var pm = FM.ExtractTypeFromArray<HNPAGEMAP>(buf, h.ibHnpm);
                    var b = new HNDataBlock
                    {
                        Index = index,
                        Buffer = buf,
                        rgibAlloc = FM.ArrayFromArray<UInt16>(buf, h.ibHnpm + Marshal.SizeOf(pm), pm.cAlloc + 1),  //+1 to get the dummy entry that gives us the size of the last one
                    };
                    blocks.Add(b);
                }
                index++;
            }
            return blocks;
        }
        private static Property ReadTableCell(List<HNDataBlock> blocks, List<SLENTRY> subnodes, RowDataBlock db, long rowOffset, TCOLDESC col)
        {   // it creates a property object with the cell content
            PCBTH pcbth = new PCBTH()
            {
                wPropId = col.wPropId,
                wPropType = col.wPropType,
            };
            byte[] data = new byte[col.cbData];   // default for data stored in the HNID/fixed size <=8
            int cellLength = CheckTypeLength(col);
            if (cellLength < 0) throw new Exception($"Unexpected property type {col.wPropType} length {col.cbData}");
            if (cellLength > 0)
            {  // fixed size types stored in the cell
                Array.Copy(db.Buffer, rowOffset + col.ibData, data, 0, cellLength);
            }
            else
            { // variable sized cells - reffered by hnid
                HNID hnid = FM.ExtractTypeFromArray<HNID>(db.Buffer, (int)rowOffset + col.ibData);
                pcbth.dwValueHnid = hnid;                // maybe overwriten
                data = GetBytesForHNID(blocks, subnodes, hnid);
            }
            return new Property(pcbth, data);
        }
        private static int CheckTypeLength(TCOLDESC col)
        {   // validates and return cell type length:
            // -1: -> invalid type/length
            //  0: -> variable cell length (ref by hnid)
            //  x: -> fixed type lenghts: 1, 2, 4 or 8  == cbData
            int result = col.cbData;
            switch (col.wPropType)
            {
                // 1 byte types
                case EpropertyType.PtypBoolean:
                    if (col.cbData != 1) result = -1;
                    break;
                // 2 byte types
                case EpropertyType.PtypInteger16:
                    if (col.cbData != 2) result = -1;
                    break;
                // 4 byte types
                case EpropertyType.PtypInteger32:
                    if (col.cbData != 4) result = -1;
                    break;
                // 8 byte types
                case EpropertyType.PtypInteger64:
                case EpropertyType.PtypFloating64:
                case EpropertyType.PtypTime:
                    // In a Table Context, time values are held in line
                    if (col.cbData != 8) result = -1;
                    break;
                // variables that are ref by a hnid
                case EpropertyType.PtypString:  // Unicode string
                case EpropertyType.PtypString8: // Multibyte string in variable encoding
                case EpropertyType.PtypBinary:
                case EpropertyType.PtypObject:
                case EpropertyType.PtypGuid:
                case EpropertyType.PtypMultipleInteger32:
                case EpropertyType.PtypMultipleString:
                case EpropertyType.PtypMultipleBinary:
                    if (col.cbData == 4) result = 0;
                    else result = -1;
                    break;
                default:
                    /// unknown property type!!!
                    result = -1;
                    break;
            }
            return result;
        }
        private static byte[] GetBytesForHNID(List<HNDataBlock> blocks, List<SLENTRY> subnodes, HNID hnid)
        {
            byte[] buf = null;

            if (hnid.hidType == EnidType.HID)
            {
                if (hnid.GetIndex(FM.srcFile.type == FileType.OST) != 0)
                {
                    buf = MapArray<byte>(blocks, hnid.HID, HidSize(blocks, hnid.HID));
                }
            }
            else if (hnid.nidType == EnidType.LTP)
            {
                int ix = subnodes.FindIndex(n => n.nid == hnid.NID.dwValue);
                buf = FM.srcFile.ReadFullDatablock(subnodes[ix].bidData);
            }
            else
                throw new Exception("Data storage style not implemented");

            return buf;
        }
        private static T[] MapArray<T>(List<HNDataBlock> blocks, HID hid, int count, int offset = 0)
        {
            var b = blocks[hid.GetBlockIndex(FM.srcFile.type == FileType.OST)];
            return FM.MapArray<T>(b.Buffer, b.rgibAlloc[hid.GetIndex(FM.srcFile.type == FileType.OST) - 1] + offset, count);
        }
        private static T MapType<T>(List<HNDataBlock> blocks, HID hid)
        {
            var b = blocks[hid.GetBlockIndex(FM.srcFile.type == FileType.OST)];
            return FM.ExtractTypeFromArray<T>(b.Buffer, b.rgibAlloc[hid.GetIndex(FM.srcFile.type == FileType.OST) - 1]);
        }
        #endregion
    }
}