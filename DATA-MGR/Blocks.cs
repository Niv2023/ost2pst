namespace ost2pst
{
    class dataBlock
    {   // used on block tree lists: SL/XX/Xblock
        public BREF bref;
        public uint cb;
        public uint cRef;
        public dataBlock(BREF Bref, uint Cb)
        {
            bref = Bref;
            cb = Cb;
            cRef = 2; // each datablock starts with 2 refs
        }
        public dataBlock() { }
    }
    struct gapBlock
    {
        public UInt64 ib;
        public uint cb;
    }
    public struct BidMapping
    {
        public UInt64 srcBid;
        public BREF outBREF;
    }

    public static class Blocks
    {
        public static PstFile pstFile;
        public static List<BidMapping> bidMappings;
        public static BREF AddDatablock(byte[] datablock, UInt64 srcBid = 0)
        {
            BREF blockBref = GetBidMapping(srcBid);
            if (blockBref.bid > 0)
            {   // srcBid already built
                pstFile.IncrementRefCount(blockBref.bid);
            }
            else
            {
                uint cb = (uint)datablock.Length;
                if (cb <= MS.MaxDatablockSize)
                {
                    blockBref = pstFile.NewDataBlock(cb);
                    WriteDataBlock(blockBref, ref datablock);
                }
                else
                {  // uncompressed OST block larger than 8192
                   // make a new x/sl tree on the output pst
                    blockBref = AddDataTree(datablock);
                }
                AddBidMapping(srcBid, blockBref);
            }
            return blockBref;
        }
        public static BREF AddDatablock(List<byte[]> datablocks, UInt64 srcBid = 0)
        {
            BREF blockBref = GetBidMapping(srcBid);
            if (blockBref.bid > 0)
            {   // srcBid already built
                pstFile.IncrementRefCount(blockBref.bid);
            }
            else
            {
                if (datablocks.Count == 1)
                {
                    blockBref = AddDatablock(datablocks[0], srcBid);
                }
                else
                {
                    blockBref = AddDataTree(datablocks);
                }
            }
            return blockBref;
        }
        public static SLENTRY CopySLentry(SLENTRY sl)
        {
            SLENTRY newSl = new SLENTRY() { nid = sl.nid };
            byte[] slDatablock;
            BREF blockBref = GetBidMapping(sl.bidData);
            if (blockBref.bid > 0)
            {   // srcBid already built
                pstFile.IncrementRefCount(blockBref.bid);
                newSl.bidData = blockBref.bid;
            }
            else
            {
                slDatablock = FM.srcFile.ReadFullDatablock(sl.bidData);
                BREF newBref = AddDatablock(slDatablock, sl.bidData);
                newSl.bidData = newBref.bid;
            }
            if (sl.bidSub > 0)
            {
                blockBref = GetBidMapping(sl.bidSub);
                if (blockBref.bid > 0)
                {   // srcBid already built
                    pstFile.IncrementRefCount(blockBref.bid);
                }
                else
                {
                    List<SLENTRY> subnodes = FM.srcFile.GetSLentries(sl.bidSub);
                    blockBref = CopySubnodes(subnodes);
                }
                newSl.bidSub = blockBref.bid;
            }
            return newSl;
        }
        public static BREF CopySubnode(UInt64 bidSub, SLENTRY? newEntry = null)
        {
            List<SLENTRY> srcSubnodes = FM.srcFile.GetSLentries(bidSub);
            return CopySubnodes(srcSubnodes, newEntry);
        }
        public static BREF CopySubnodes(List<SLENTRY> srcSubnodes,SLENTRY? newEntry = null)
        {
            List<SLENTRY> destSubnodes = new List<SLENTRY>();
            foreach (SLENTRY s in srcSubnodes)
            {
                if (s.nid == newEntry?.nid)
                {
                    destSubnodes.Add((SLENTRY)newEntry);
                }
                else
                {
                    destSubnodes.Add(CopySLentry(s));
                }
                FM.UpdateSubnodeNIDheaderIndex((UInt32)s.nid);
            }
            return AddSLentries(destSubnodes);
        }
        public static BREF CopyDatablock(UInt64 srcBid)
        {
            byte[] data = FM.srcFile.ReadFullDatablock(srcBid);
            return AddDatablock(data, srcBid);
        }
        private static void AddBidMapping(UInt64 srcBid, BREF bref)
        {
            if (srcBid > 0)
            {
                int index = bidMappings.FindIndex(x => x.srcBid == srcBid);
                if (index >= 0) throw new Exception($"ERROR MAPPING BID: source bid {srcBid} already mapped");
                BidMapping bidMap = new BidMapping() { srcBid = srcBid, outBREF = bref };
                bidMappings.Add(bidMap);
            }
        }
        private static BREF GetBidMapping(UInt64 srcBid)
        {
            BREF blockBref = new();
            int index = bidMappings.FindIndex(x => x.srcBid == srcBid);
            if (index > 0) blockBref = bidMappings[index].outBREF;
            return blockBref;
        }
        public static BREF AddDataTree(List<byte[]> datablocks)
        {
            uint nrDatablocks = (uint)datablocks.Count;
            BREF dataBREF = new BREF();
            uint IcbTotal = 0;
            uint cbEntry = 0;
            bool XXblock = nrDatablocks > MS.MaxXblockEntries;
            dataBlock XXdatablock = new dataBlock();
            List<dataBlock> Xdatablocks = new List<dataBlock>();
            List<dataBlock> datablockLeaves = new List<dataBlock>();
            if (XXblock)
            {
                XXdatablock = AddXXblock(nrDatablocks);
                cbEntry = XXdatablock.cb;
            }
            for (int i = 0; i < nrDatablocks; i++)
            {
                uint blockSize = (uint)datablocks[(int)i].Length;
                if (blockSize > MS.MaxBlockSize)
                {
                    throw new Exception($"block too large");
                }
                IcbTotal += blockSize;
                if (i % MS.MaxXblockEntries == 0)
                {
                    if (i == 0 & !XXblock)
                    {
                        Xdatablocks.Add(AddXblock((uint)i, nrDatablocks));
                        cbEntry = Xdatablocks[0].cb;

                    }
                    else
                    {
                        Xdatablocks.Add(AddXblock((uint)i, nrDatablocks));
                    }
                }
                datablockLeaves.Add(new dataBlock(AddDatablock(datablocks[i]), blockSize));
            }
            dataBREF = UpdateXBlocks(Xdatablocks, datablockLeaves); // returns first XBlock BREF
            if (XXblock)
            {
                dataBREF = UpdateXXBlock(XXdatablock, Xdatablocks, IcbTotal); // returns the XXBlock BREF
            }
            return dataBREF;
        }
        private static BREF AddDataTree(byte[] datablock)
        {
            List<byte[]> datablocks = BlockList(datablock);
            return AddDataTree(datablocks);
        }
        private static dataBlock AddXXblock(uint nrBlocks)
        {
            uint nrXblocks = nrBlocks / MS.MaxXblockEntries;
            if (nrBlocks % MS.MaxXblockEntries != 0) { nrXblocks++; }
            uint xxBlockSize = 8 + nrXblocks * 8;
            return new dataBlock(pstFile.NewDataBlock(xxBlockSize, true), xxBlockSize);
        }
        private static dataBlock AddXblock(uint curBlock, uint nrBlocks)
        {
            uint nrEntries = nrBlocks - curBlock;
            if (nrEntries > MS.MaxXblockEntries) nrEntries = MS.MaxXblockEntries;
            uint xBlockSize = 8 + nrEntries * 8;
            return new dataBlock(pstFile.NewDataBlock(xBlockSize, true), xBlockSize);  // assigns a new bid
        }
        private static BREF UpdateXXBlock(dataBlock xxBlock, List<dataBlock> xBlocks, uint icbTotal)
        {
            byte[] xxData = new byte[xxBlock.cb];
            xxData[0] = 1; // btype
            xxData[1] = 2; // xx level
            UInt16 cEnt = (UInt16)xBlocks.Count;
            FM.TypeToArray<UInt16>(ref xxData, cEnt, 2);
            FM.TypeToArray<UInt64>(ref xxData, icbTotal, 4);
            for (int i = 0; i < cEnt; i++)
            {
                FM.TypeToArray<UInt64>(ref xxData, xBlocks[i].bref.bid, 8 + i * 8);
            }
            WriteDataBlock(xxBlock.bref, ref xxData);  // false = internal datablock
            return xxBlock.bref;
        }
        private static BREF UpdateXBlocks(List<dataBlock> xBlocks, List<dataBlock> blocks)
        {
            int nrXblock = xBlocks.Count;
            for (int i = 0; i < nrXblock; i++)
            {
                UInt32 xIcbTotal = 0;
                uint xBlockSize = xBlocks[i].cb;
                byte[] xData = new byte[xBlockSize];
                xData[0] = 1; // btype
                xData[1] = 1; // x level
                UInt16 cEnt = (UInt16)((xBlocks[i].cb - 8) / 8);
                FM.TypeToArray<UInt16>(ref xData, cEnt, 2);
                for (int j = 0; j < cEnt; j++)
                {
                    int blockIx = i * (int)MS.MaxXblockEntries + j;
                    xIcbTotal += blocks[blockIx].cb;
                    FM.TypeToArray<UInt64>(ref xData, blocks[blockIx].bref.bid, 8 + j * 8);
                }
                FM.TypeToArray<UInt32>(ref xData, xIcbTotal, 4);
                WriteDataBlock(xBlocks[i].bref, ref xData);  // false = internal datablock
            }
            return xBlocks[0].bref;
        }
        private static List<byte[]> BlockList(byte[] datablock)
        {
            List<byte[]> list = new List<byte[]>();
            int nrDatablocks = (int)(datablock.Length / MS.MaxDatablockSize);
            if (datablock.Length % MS.MaxDatablockSize != 0) { nrDatablocks++; }
            int lastBlockSize = (int)(datablock.Length - MS.MaxDatablockSize * (nrDatablocks - 1));
            for (int i = 0; i < nrDatablocks - 1; i++)
            {
                byte[] block = new byte[MS.MaxDatablockSize];
                Array.ConstrainedCopy(datablock, (int)(i * MS.MaxDatablockSize), block, 0, block.Length);
                list.Add(block);
            }
            byte[] lastBlock = new byte[lastBlockSize];
            Array.ConstrainedCopy(datablock, (int)((nrDatablocks - 1) * MS.MaxDatablockSize), lastBlock, 0, lastBlockSize);
            list.Add(lastBlock);
            return list;
        }
        private unsafe static void WriteDataBlock(BREF bref, ref byte[] data)
        {
            int totalSize = 0; int trailer = 0;
            if ((bref.bid & 0x02) == 0) { PST.CryptPermuteEncode(data); }   // bid external block
            uint crc = PST.ComputeCRC(data);
            uint cb = (uint)data.Length;
            totalSize = (int)PST.BlockTotalSize(cb);
            trailer = totalSize - sizeof(BLOCKTRAILER);
            Array.Resize(ref data, totalSize);
            BLOCKTRAILER blockTrailer = new BLOCKTRAILER()
            {
                cb = (ushort)cb,
                wSig = PST.ComputeSig(bref.ib, bref.bid),
                bid = bref.bid,
                dwCRC = crc
            };
            FM.ExtractArrayFromType<BLOCKTRAILER>(ref data, blockTrailer, sizeof(BLOCKTRAILER), 0, trailer);
            pstFile.stream.Position = (long)bref.ib;
            pstFile.stream.Write(data, 0, data.Length);
        }
        #region SL/SI blocks
        public static BREF AddSLentries(List<SLENTRY> slEntries)
        {
            BREF SLbref = new BREF();
            int nrSLentries = slEntries.Count;
            if (nrSLentries > 0)
            {
                if (nrSLentries > MS.MaxSLblockEntries)
                {
                    // add SI block
                    return AddSIentries(slEntries);
                }
                uint cb = (uint)nrSLentries * 24 + 8;
                SLbref = pstFile.NewDataBlock(cb, true);
                byte[] SLData = new byte[cb];
                SLData[0] = 2; // SLtype
                SLData[1] = 0; // SL level
                FM.TypeToArray<UInt16>(ref SLData, (UInt16)nrSLentries, 2);
                for (int i = 0; i < nrSLentries; i++)
                {
                    FM.TypeToArray<UInt64>(ref SLData, slEntries[i].nid, 8 + i * 24);
                    FM.TypeToArray<UInt64>(ref SLData, slEntries[i].bidData, 8 + 8 + i * 24);
                    FM.TypeToArray<UInt64>(ref SLData, slEntries[i].bidSub, 8 + 16 + i * 24);
                }
                WriteDataBlock(SLbref, ref SLData);
            }
            return SLbref;
        }
        private static BREF AddSIentries(List<SLENTRY> subNode)
        {
            int nrSIEntries = subNode.Count / (int)MS.MaxSLblockEntries;
            if (nrSIEntries % (int)MS.MaxSLblockEntries != 0) { nrSIEntries++; }
            uint cb = (uint)nrSIEntries * 16 + 8;
            if (cb > MS.MaxDatablockSize) { throw new Exception($"SL/SI subnode overflow. Cannot support {nrSIEntries} entries"); }
            BREF SIbref = pstFile.NewDataBlock(cb, true);
            byte[] SIData = new byte[cb];
            SIData[0] = 2; // SItype
            SIData[1] = 1; // SL level
            FM.TypeToArray<UInt16>(ref SIData, (UInt16)nrSIEntries, 2);
            for (int i = 0; i < nrSIEntries; i++)
            {
                int nrSLentries = (int)MS.MaxSLblockEntries;
                if (i == nrSIEntries - 1) { nrSLentries = subNode.Count - (nrSIEntries - 1) * (int)MS.MaxSLblockEntries; }
                List<SLENTRY> SLnodes = subNode.GetRange(i * (int)MS.MaxSLblockEntries, (int)nrSLentries);
                SLENTRY firstNode = subNode[i * (int)MS.MaxSLblockEntries];
                BREF slBref = AddSLentries(SLnodes);
                SIENTRY SIEntry = new SIENTRY()
                {
                    nid = firstNode.nid,   /// need to guarantee nid uniqueness within the sl tree
                    bid = slBref.bid,
                };
                FM.TypeToArray<SIENTRY>(ref SIData, SIEntry, 8 + i * 16);
            }
            WriteDataBlock(SIbref, ref SIData);
            return SIbref;
        }
        #endregion
    }
}
