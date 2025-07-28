/*  PST data File Manager (MS-PST v20221115)
 *  Manages the layer 1 PST data structures;
 *  - PST hearder
 *  - DList
 *  - AMap allocation
 *      -> adds dummy PMap, FMap and FPMap for comkeep data/node allocation and file integrity
 *  Notes:
 *  - Although the DList is kept updated, in this initial version a new AMap will be created whenever
 *    the previous AMap cannot support a data allocation request
 *  - Since data blocks are assigned in sequence (never released), all AMap free slots are at the end
 */

namespace ost2pst
{
    /* 
    * PstFile class
    * - Manages the output pst file
    * - output file is a PST UNICODE
    */

    public class PstFile
    {
        public string filename;
        public FileStream stream;
        public FileHeader header;
        public List<NBTENTRY> NBTs;    // list of NBT entries (OST and PST)
        public List<BBTENTRY> BBTs;    // list of BBT entries (PST)
        public static DListPage DListPage;
        public static List<AMap> AMapList;
        private List<gapBlock> gapBlocks;            // blocks allocated for page block allignment. AMap will be released

        public PstFile(string fn)
        {
            filename = fn;
            try
            {
                stream = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite);
                Blocks.pstFile = this;
                Blocks.bidMappings = new List<BidMapping>();
                NBTs = new List<NBTENTRY>();
                BBTs = new List<BBTENTRY>();
                DListPage = new DListPage();
                AMapList = new List<AMap>();
                BuildPSTheader();
                UpdatePSTheader(); // file lenght... crc...
                BuildDListPage();
            }
            catch (Exception ex)
            {
                stream?.Close();
                throw new Exception($"ERROR CREATING PST: {ex.Message}");
            }
        }
        public void Close()
        {
            stream?.Close();
        }

        public void CopyNDB(NBTENTRY nbt)
        {
            BREF dataBREF = Blocks.CopyDatablock(nbt.bidData);
            UInt64 bidSub = nbt.bidSub;
            if (bidSub > 0)
            {
                BREF subBREF = Blocks.CopySubnode(nbt.bidSub);
                bidSub = subBREF.bid;
            }
            NBTENTRY newNbt = new NBTENTRY()
            {
                nid = nbt.nid,
                bidData = dataBREF.bid,
                nidParent = nbt.nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(newNbt);
        }
        public void AddNDB(NBTENTRY nbt, List<byte[]> datablocks, UInt64 newSubnode)
        {
            BREF dataBREF = Blocks.AddDatablock(datablocks, nbt.bidData);
            NBTENTRY newNbt = new NBTENTRY()
            {
                nid = nbt.nid,
                bidData = dataBREF.bid,
                nidParent = nbt.nidParent,
                bidSub = newSubnode,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(newNbt);
        }
        public UInt64 AddNDBdata(NID nid, byte[] datablock, UInt64 bidSub, UInt32 nidParent)
        {
            BREF dataBREF = Blocks.AddDatablock(datablock);
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = nid,
                bidData = dataBREF.bid,
                nidParent = nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(nbt);
            return dataBREF.bid;
        }
        public UInt64 AddNDBdata(NID nid, List<byte[]> datablocks,  UInt64 bidSub, UInt32 nidParent)
        {
            BREF dataBREF = new();
            if (datablocks.Count == 1)
            {
                dataBREF = Blocks.AddDatablock(datablocks[0]);
            }
            else
            {
                dataBREF = Blocks.AddDataTree(datablocks);
            }
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = nid,
                bidData = dataBREF.bid,
                nidParent = nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(nbt);
            return dataBREF.bid;
        }
        public NBTENTRY AddNDBdataNbt(NID nid, List<byte[]> datablocks, UInt64 bidSub, UInt32 nidParent)
        {
            BREF dataBREF = new();
            if (datablocks.Count == 1)
            {
                dataBREF = Blocks.AddDatablock(datablocks[0]);
            }
            else
            {
                dataBREF = Blocks.AddDataTree(datablocks);
            }
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = nid,
                bidData = dataBREF.bid,
                nidParent = nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(nbt);
            return nbt;
        }
        public NBTENTRY AddSubnodeNDBdataNbt(NID nid, List<byte[]> datablocks, UInt64 bidSub, UInt32 nidParent)
        {
            BREF dataBREF = new();
            if (datablocks.Count == 1)
            {
                dataBREF = Blocks.AddDatablock(datablocks[0]);
            }
            else
            {
                dataBREF = Blocks.AddDataTree(datablocks);
            }
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = nid,
                bidData = dataBREF.bid,
                nidParent = nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            return nbt;
        }
        public void AddNDBdata(NID nid, UInt64 bidData,  UInt32 nidParent, UInt64 bidSub)
        {
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = nid,
                bidData = bidData,
                nidParent = nidParent,
                bidSub = bidSub,
                dwPad = 0,
                dwPadding = 0
            };
            NBTs.Add(nbt);
        }
        public void IncrementRefCount(UInt64 bid)
        {
            if (bid > 0)
            {
                int ix = BBTs.FindIndex(x => x.BREF.bid == bid);
                BBTENTRY bbt = BBTs[ix];
                bbt.cRef++;
                BBTs[ix] = bbt;
            }
        }
        private unsafe void UpdatePSTheader()
        {
            ulong pstSize = (ulong)stream.Length;
            header.root.ibFileEof = pstSize;
            header.dwCRCPartial = GetPartialCRC();
            header.dwCRCFull = GetFullCRC();
            stream.Position = 0;
            FM.WriteType<FileHeader>(stream, header);   // header on the begining of the pst file (offset = 0)
        }
        private void BuildDListPage()
        {
            DListPage.pageTrailer = GetPageTrailer<DListPage>(Eptype.ptypeDL, MS.dListOffset, DListPage);
            FM.WriteType<DListPage>(stream, DListPage, MS.dListOffset);
        }
        /// <summary>
        /// Adds PageBlock (512 bytes)
        /// refer to MS_PST 2.2.2.7.7.1 BTPAGE
        /// - page blocks are not added in the BBT tree
        /// - pages should be aligned on 0x200 offset address
        ///     . this is not mentioned in the MS_PST
        ///     . outlook give an error message
        /// </summary>
        public unsafe BREF AddPage(List<BTENTRY> entries, Eptype eptype, int level)
        {
            BTPAGE btPage = new BTPAGE()
            {
                cEnt = (byte)entries.Count,
                cbEnt = MS.BTcbEnt,
                cEntMax = (byte)(MS.BTPageEntryBytes / MS.BTcbEnt),
                cLevel = (byte)level,
                dwPadding = 0
            };
            for (int i = 0; i < entries.Count; i++)
            {
                FM.TypeToBuffer<BTENTRY>(btPage.rgentries, MS.BTPageEntryBytes, entries[i], i * btPage.cbEnt);

            }
            BREF pageBREF = NewPageBlock();
            btPage.pageTrailer = GetPageTrailer<BTPAGE>(eptype, pageBREF.ib, btPage, pageBREF.bid);
            FM.WriteType<BTPAGE>(stream, btPage, pageBREF.ib);
            return pageBREF;
        }
        public unsafe BREF AddPage(List<BBTENTRY> entries)
        {
            BTPAGE btPage = new BTPAGE()
            {
                cEnt = (byte)entries.Count,
                cbEnt = MS.BBTcbEnt,
                cEntMax = (byte)(MS.BTPageEntryBytes / MS.BBTcbEnt),
                cLevel = 0,                 // BBTENTRY on level 0
                dwPadding = 0
            };
            for (int i = 0; i < entries.Count; i++)
            {
                FM.TypeToBuffer<BBTENTRY>(btPage.rgentries, MS.BTPageEntryBytes, entries[i], i * btPage.cbEnt);
            }
            BREF pageBREF = NewPageBlock();
            btPage.pageTrailer = GetPageTrailer<BTPAGE>(Eptype.ptypeBBT, pageBREF.ib, btPage, pageBREF.bid);
            FM.WriteType<BTPAGE>(stream, btPage, pageBREF.ib);
            return pageBREF;
        }
        public unsafe BREF AddPage(List<NBTENTRY> entries)
        {
            BTPAGE btPage = new BTPAGE()
            {
                cEnt = (byte)entries.Count,
                cbEnt = MS.NBTcbEnt,
                cEntMax = (byte)(MS.BTPageEntryBytes / MS.NBTcbEnt),
                cLevel = 0,                 // BBTENTRY on level 0
                dwPadding = 0
            };
            for (int i = 0; i < entries.Count; i++)
            {
                FM.TypeToBuffer<NBTENTRY>(btPage.rgentries, MS.BTPageEntryBytes, entries[i], i * btPage.cbEnt);
            }
            BREF pageBREF = NewPageBlock();
            btPage.pageTrailer = GetPageTrailer<BTPAGE>(Eptype.ptypeNBT, pageBREF.ib, btPage, pageBREF.bid);
            FM.WriteType<BTPAGE>(stream, btPage, pageBREF.ib);
            return pageBREF;
        }
        public void FinalizePSTdata()
        {
            ReleaseGapBlocks();
            UpdateDListPage();
            UpdatePSTheader();
        }
        public BREF NewDataBlock(uint cb, bool isInternal = false)
        {
            if (cb > MS.MaxDatablockSize) { throw new Exception("DATA BLOCK ERROR: exceeds 8176 bytes"); }
            BBTENTRY bbt = new BBTENTRY();
            uint totalSize = PST.BlockTotalSize(cb);
            bbt.BREF.ib = GetNewBlockOffset(totalSize);
            bbt.BREF.bid = GetNextDataBid();
            bbt.cb = (ushort)cb;
            bbt.cRef = 2;
            if (isInternal)
            {   // set "b" bit in the bid
                bbt.BREF.bid = bbt.BREF.bid | 2;
            }
            BBTs.Add(bbt);
            return bbt.BREF;
        }
        private BREF NewPageBlock()
        {
            return new BREF()
            {
                ib = GetNewBlockOffset(MS.NodePageBlockSize, true), // pages must be aligned 0x200
                bid = GetNextPageBid()
            };
        }
        private uint GetPartialCRC()
        {
            int crcPartialLen = 471;
            byte[] array = new byte[crcPartialLen];
            FM.ExtractArrayFromType<FileHeader>(ref array, header, crcPartialLen, 8);  // 471 bytes from hearder pos 8
            uint crc = PST.ComputeCRC(array, crcPartialLen);
            return crc;
        }
        private uint GetFullCRC()
        {
            int crcFullLen = 516;
            byte[] buffer = new byte[crcFullLen];
            FM.ExtractArrayFromType<FileHeader>(ref buffer, header, crcFullLen, 8);
            uint crc = PST.ComputeCRC(buffer, crcFullLen);
            return crc;
        }
        private void BuildPSTheader()
        {
            InitializeHeader();
            FM.WriteType<FileHeader>(stream, header);   // header on the begining of the pst file (offset = 0)
            byte[] data = new byte[16332];              // nr of bytes between the hearder and offset 0x4200
            stream.Write(data, 0, data.Length);
        }
        private unsafe void InitializeHeader()
        {
            header = new FileHeader()
            {
                dwMagic = MS.Magic,
                dwCRCPartial = 0, // Tbd
                wMagicClient = MS.MagicClient,
                wVer = MS.minUnicodeVer,
                wVerClient = MS.VerClient,
                bPlatformCreate = 1,
                bPlatformAccess = 1,
                dwReserved1 = 0,
                dwReserved2 = 0,
                bidUnused = 0,          // 8 bytes unicode
                bidNextP = 1,           // 8 bytes unicode
                dwUnique = 1,           // 4 bytes
                                        //   rgnid[32],         // 32 x NID index table (128 bytes)
                qwUnused = 0,           // 8 bytes unicode
                root = new Root()            // 72 bytes unicode
                {
                    dwReserved = 0,                                 // 4 bytes unicode
                    ibFileEof = 0,                                  // 8 bytes unicode = size of the PST file
                    ibAMapLast = 0,                                 // 8 bytes -> absolute file offset to the last AMap page in PST file
                    cbAMapFree = 0,                                 // 8 bytes unicode - total free space in all AMaps combined.
                    cbPMapFree = 0,                                 // 8 bytes unicode - deprecated 
                    BREFNBT = new BREF(),                           // 16 bytes unicode -> BREF struct root page Node BTH (NBT)
                    BREFBBT = new BREF(),                           // 16 bytes unicode -> BREF struct root page Block BTH (BBT)
                    fAMapValid = 0x02,                              // 1 byte VALID_AMAP2 - the AMaps are valid
                    bReserved = 0,                                  // 1 byte
                    wReserved = 0                                   // 2 bytes
                },
                dwAlign = 0,                                        // 4 bytes unicode
                /*
                fixed Byte rgbFM[128];
                fixed Byte rgbFP[128];
                */
                bSentinel = 0x80,
                bCryptMethod = EbCryptMethod.NDB_CRYPT_PERMUTE,     // PST will be created with CRYPT_PERMUTE
                rgbReserved = 0,                                    // 2 bytes
                bidNextB = 4,                                       // 8 bytes unicode
                dwCRCFull = 0,                                      // 4 bytes - 516 bytes CRC from dwMagicClient to BidNextB
                rgbReserved2 = 0                                    // 3 bytes + 1 bytr bReserved
                /*
                fixed Byte rgbReserved3[32];  // 32 bytes
                */
            };
            for (int i = 0; i < 128; i++)
            {
                header.rgbFM[i] = 0xff;
                header.rgbFP[i] = 0xff;
            }
            for (int i = 0; i < 32; i++)
            {
                header.rgnid[i] = 0x400;
                if ((EnidType)i == EnidType.SEARCH_FOLDER) { header.rgnid[i] = 0x4000; }
                if ((EnidType)i == EnidType.NORMAL_MESSAGE) { header.rgnid[i] = 0x10000; }
                if ((EnidType)i == EnidType.ASSOC_MESSAGE) { header.rgnid[i] = 0x8000; }
            }
        }
        #region PST AMap management - only used for output PST
        private UInt64 GetNewBlockOffset(uint size, bool align512 = false)
        {   // for a new single data block allocation
            UInt64 offset = 0;
            if (size > MS.MaxBlockSize) { throw new Exception($"Data block request for {size} bytes failed. Max is {MS.MaxBlockSize} "); }
            if (SpaceAvailableOnLastAMap() >= size)
            {
                // all free space is sequencial at end of the Data section
                offset = FindFirstFreeBlock();
                if (align512)
                {   // page data blocks needs to be alligned on 0x0200 offset address
                    // this is not in the MS_PST, but outlook gives an error
                    if (offset % 0x200 != 0)
                    {
                        uint gap = (uint)(0x200 - (offset % 0x200));
                        AllocateDataOnLastAMap(offset, gap);
                        AddGAPentry(offset, gap);                   // will be released before closing pst
                        offset = FindFirstFreeBlock();
                    }
                }
                AllocateDataOnLastAMap(offset, size);
            }
            else
            {
                GetNewAMap();
                return GetNewBlockOffset(size, align512);
            }
            return offset;
        }
        private void ReleaseGapBlocks()
        {
            if (gapBlocks == null) return;
            for (int i = 0; i < gapBlocks.Count; i++)
            {
                ReleaseBlock(gapBlocks[i].ib, gapBlocks[i].cb);
            }
        }
        unsafe private void ReleaseBlock(UInt64 offset, uint size)
        {   // size MUST be a multiple of 64
            int index = AMapList.FindIndex(x => x.offset > offset);
            if (index < 0)
            {   // it should be in the last AMap
                index = AMapList.Count - 1;
            }
            else
            {   // previous AMap
                index -= 1;
            }
            UInt64 aMapOffset = AMapList[index].offset;
            AMapPage aMapPage = FM.ReadType<AMapPage>(stream, aMapOffset);
            int firstBit = (int)((offset - aMapOffset) / 64); // block should start on multiple of 64
            int lastBit = firstBit + (int)(size / 64) - 1;
            for (int i = firstBit; i <= lastBit; i++)
            {
                PST.ResetBitInArray(ref aMapPage, i);
            }
            AMapList[AMapList.Count - 1].nrFreeSlots += size / 64;
            header.root.cbAMapFree += size;
            UpdateAMap(ref aMapPage, aMapOffset);
        }
        private void UpdateDListPage()
        {   // ------------> CRC should be done on the number of entries
            AMapList.Sort(); // by free slots
            int nEntries = MS.AMapPageEntryBytes;
            byte[] array = new byte[nEntries];
            FM.ExtractArrayFromType(ref array, DListPage, nEntries);
            //Map.ExtractBytes(ref buffer, DListPage, nEntries);
            DListPage.pageTrailer.dwCRC = PST.ComputeCRC(array, nEntries);
            FM.WriteType<DListPage>(stream, DListPage, MS.dListOffset);
        }
        private uint SpaceAvailableOnLastAMap()
        {
            if (AMapList.Count == 0) { return 0; }
            return AMapList[AMapList.Count - 1].freeSpace;
        }
        private unsafe UInt64 FindFirstFreeBlock()
        {
            UInt64 aMapOffset = header.root.ibAMapLast;
            uint bitIx = 0;
            AMapPage aMap = FM.ReadType<AMapPage>(stream, aMapOffset);
            for (int i = 0; i < MS.AMapPageEntryBytes; i++)
            {
                if (aMap.rgbAMapBits[i] == 0xFF)
                {
                    bitIx += 8;
                }
                else
                {
                    byte val = aMap.rgbAMapBits[i];
                    for (int j = 0; j < 8; j++)
                    {
                        if (PST.GetBitInByte(j, val) == 0) { break; }
                        bitIx++;
                    }
                }
            }
            return (UInt64)(aMapOffset + (bitIx * 64));
        }
        private unsafe void AllocateDataOnLastAMap(UInt64 offset, uint size)
        {   // size MUST be a multiple of 64
            UInt64 aMapOffset = header.root.ibAMapLast;
            AMapPage aMapPage = FM.ReadType<AMapPage>(stream, aMapOffset);
            int firstBit = (int)((offset - aMapOffset) / 64); // block should start on multiple of 64
            int lastBit = firstBit + (int)(size / 64) - 1;
            for (int i = firstBit; i <= lastBit; i++)
            {
                PST.SetBitInArray(ref aMapPage, i);
            }
            AMapList[AMapList.Count - 1].nrFreeSlots -= size / 64;
            header.root.cbAMapFree -= size;
            UpdateAMap(ref aMapPage, aMapOffset);
        }
        private void UpdateAMap(ref AMapPage aMapPage, UInt64 offset) // off
        {
            int nrBytes = MS.AMapPageEntryBytes;
            byte[] array = new byte[nrBytes];
            FM.ExtractArrayFromType<AMapPage>(ref array, aMapPage, nrBytes);
            aMapPage.pageTrailer.dwCRC = PST.ComputeCRC(array, nrBytes);
            FM.WriteType<AMapPage>(stream, aMapPage, offset);
        }
        private void AddGAPentry(UInt64 offset, uint lenght)
        {
            if (gapBlocks == null) { gapBlocks = new List<gapBlock>(); }
            gapBlock gap = new gapBlock()
            {
                ib = offset,
                cb = lenght
            };
            gapBlocks.Add(gap);
        }
        private unsafe void GetNewAMap()
        {   // called when the last AMap cannot allocate a new data block
            uint pageTrail = MS.AMapAllocationSize - 512; // total bytes after AMapPage
            if (header.root.ibAMapLast == 0)
            {
                header.root.ibAMapLast = MS.firstAMapOffset;
            }
            else
            {
                header.root.ibAMapLast += MS.AMapAllocationSize;
            }
            AMapPage aMapPage = new AMapPage();
            PMapPage pMapPage = new PMapPage();
            FMapPage fMapPage = new FMapPage();
            FPMapPage fpMapPage = new FPMapPage();
            for (int i = 0; i < MS.AMapPageEntryBytes; i++)
            {
                pMapPage.rgbPMapBits[i] = 0xFF;
                fMapPage.rgbFMapBits[i] = 0xFF;
                fpMapPage.rgbFPMapBits[i] = 0xFF;
            }
            GetMapPages(ref aMapPage, ref pMapPage, ref fMapPage, ref fpMapPage);
            WritePage<AMapPage>(aMapPage, false, header.root.ibAMapLast);
            if (pMapPage.pageTrailer.ptype == Eptype.ptypePMap)
            {
                pageTrail -= 512;
                WritePage<PMapPage>(pMapPage);
            }

            if (fMapPage.pageTrailer.ptype == Eptype.ptypeFMap)
            {
                pageTrail -= 512;
                WritePage<FMapPage>(fMapPage);
            }
            if (fpMapPage.pageTrailer.ptype == Eptype.ptypeFPMap)
            {
                pageTrail -= 512;
                WritePage<FPMapPage>(fpMapPage);
            }
            byte[] buffer = new byte[pageTrail];
            stream.Write(buffer, 0, (int)pageTrail);
            header.root.cbAMapFree += AMapList[AMapList.Count - 1].freeSpace;
        }
        // new AMap page created and initialized
        // if applicaple, other (dummy) pages will be added
        private void GetMapPages(ref AMapPage aMapPage, ref PMapPage pMapPage, ref FMapPage fMapPage, ref FPMapPage fpMapPage)
        {
            bool pmap = false; bool fmap = false; bool fpmap = false;
            int usedSlots = 8 + CheckAdditionalMaps(ref pmap, ref fmap, ref fpmap);
            for (int i = 0; i < usedSlots; i++) { PST.SetBitInArray(ref aMapPage, i); }
            AMap aMap = new AMap
            {
                index = AMapList.Count,
                offset = header.root.ibAMapLast,
                nrFreeSlots = (uint)(MS.AMapPageEntryBytes * 8 - usedSlots)  // amap + fmap slots
            };
            AMapList.Add(aMap);
            UInt64 offset = header.root.ibAMapLast;
            aMapPage.pageTrailer = GetPageTrailer<AMapPage>(Eptype.ptypeAMap, offset, aMapPage);
            if (pmap)
            {
                offset += 512; //after AMap
                pMapPage.pageTrailer = GetPageTrailer<PMapPage>(Eptype.ptypePMap, offset, pMapPage);
            }
            if (fmap)
            {
                offset += 512; //after PMap
                fMapPage.pageTrailer = GetPageTrailer<FMapPage>(Eptype.ptypeFMap, offset, fMapPage);
            }
            if (fpmap)
            {
                offset += 512; //after fMap
                fpMapPage.pageTrailer = GetPageTrailer<FPMapPage>(Eptype.ptypeFPMap, offset, fpMapPage);
            }
        }
        private PAGETRAILER GetPageTrailer<T>(Eptype eptype, UInt64 offset, T pageData, UInt64 bid = 0)
        {   // MS_PST section 2.2.2.7.1	PAGETRAILER
            PAGETRAILER pageTrailer = new PAGETRAILER()
            {
                ptype = eptype,
                ptypeRepeat = eptype,
                wSig = 0,       // for ptypeAMap, FMap, PMap, FPMap
                bid = offset    // 
            };
            if (eptype == Eptype.ptypeBBT || eptype == Eptype.ptypeNBT || eptype == Eptype.ptypeDL)
            {
                if (bid == 0)
                {
                    pageTrailer.bid = GetNextPageBid();
                }
                else
                {
                    pageTrailer.bid = bid;
                }
                pageTrailer.wSig = PST.ComputeSig(offset, pageTrailer.bid);
            }
            int dataBytes = MS.AMapPageEntryBytes;
            byte[] array = new byte[dataBytes];
            FM.ExtractArrayFromType<T>(ref array, pageData, dataBytes);
            pageTrailer.dwCRC = PST.ComputeCRC(array, dataBytes);
            return pageTrailer;
        }
        private void WritePage<T>(T page, bool computeCRC = false, UInt64 offset = 0)
        {
            if (offset > 0) { stream.Position = (long)offset; }
            if (computeCRC)
            {
                int nrBytes = MS.AMapPageEntryBytes;  // page size (512) - trailer (16) = 496
                                                      //               byte[] buffer = Map.ExtractBytes<T>(page, nrBytes);
                byte[] array = new byte[0];
                FM.ExtractArrayFromType<T>(ref array, page, nrBytes);
                uint crc = PST.ComputeCRC(array, nrBytes);
                AMapPage aPage = (AMapPage)(object)page;
                aPage.pageTrailer.dwCRC = crc;
            }
            FM.WriteType(stream, page);
        }
        private UInt64 GetNextPageBid()
        { // 2.6.1.1.4	Creating a Page
            ulong nextPageBid = header.bidNextP;
            //header.bidNextB += 4;        // to comply with the spec!!! but don't think necessary
            header.bidNextP++;
            return nextPageBid;
        }
        private UInt64 GetNextDataBid()
        {
            // ulong nextBix = (pstHeader.bidNextB >> 2) + 1;
            // pstHeader.bidNextB = nextBix << 2;
            ulong nextBid = header.bidNextB;
            header.bidNextB += 4;
            return nextBid;
        }
        private int CheckAdditionalMaps(ref bool pmap, ref bool fmap, ref bool fpmap)
        {   // each page occupies 4 slots (512 bytes)
            int nrSlots = 0; //
            int aMapCount = AMapList.Count;
            if (aMapCount % 8 == 0) { pmap = true; nrSlots += 8; }  // PMap every 8 AMaps
            if (aMapCount >= 128)
            {
                if ((aMapCount - 128) % 496 == 0) { fmap = true; nrSlots += 8; }
                if (aMapCount >= 8192)
                {
                    if ((aMapCount - 8192) % 31744 == 0) { fpmap = true; nrSlots += 8; } // 31744 = 496 * 8 * 8
                }
            }
            return nrSlots;
        }
        #endregion
    }
}