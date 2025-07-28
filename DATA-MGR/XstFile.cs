/*  
 *  Class for reading OST/PST data structures (MS-PST v20221115)
 *  
 */

using System.IO.Compression;

namespace ost2pst
{
    /* 
* XstFile class
* - source file can be OST or PST (UNICODE) file
* - output file is a PST UNICODE
*/

    public struct BBT
    {
        public EbType type;
        public BREF BREF;
        public UInt16 cbStored;
        public UInt16 cbInflated;
        public UInt16 cRef;
        public UInt16 cReadRef;
        public UInt32 dwPadding;
        public int cLevel;
        public int cEnt;
        public UInt32 IcbTotal;
    }

    public class XstFile
    {
        public string filename;
        public FileStream stream;
        public FileHeader header;
        public FileType type;
        public List<NBTENTRY> NBTs;    // list of NBT entries (OST === PST)
        public List<BBT> BBTs;        // list of BBT entries (combined OST/PST entries)
        private int pageEntries;

        public XstFile(string fn)
        {
            filename = fn;
            try
            {
                stream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                header = FM.ReadType<FileHeader>(stream);
                if (header.dwMagic != MS.Magic) { throw new Exception($"{fn} is not a PST/OST file (wrong magic key)"); }
                if (header.wVer < MS.minUnicodeVer) { throw new Exception($"{fn} is not an UNICODE PST/OST file"); }
                NBTs = new List<NBTENTRY>();
                BBTs = new List<BBT>();
                if (header.wVer == MS.OSTversion)
                {
                    type = FileType.OST;
                    pageEntries = MS.BTPAGEEntryBytesOST;
                    ReadBTtreeOST(header.root.BREFNBT.ib);
                    ReadBTtreeOST(header.root.BREFBBT.ib);
                }
                else
                {
                    type = FileType.PST;
                    pageEntries = MS.BTPageEntryBytes;
                    ReadBTtree(header.root.BREFNBT.ib);
                    ReadBTtree(header.root.BREFBBT.ib);
                }
            }
            catch (Exception ex)
            {
                stream?.Close();
                throw new Exception($"ERROR OPENNING OST/PST: {ex.Message}");
            }
        }
        public void Close()
        {
            stream?.Close();
        }
        private void ReadBTtree(ulong offset)
        {
            try
            {
                var page = FM.ReadType<BTPAGE>(stream, offset);
                for (int i = 0; i < page.cEnt; i++)
                {
                    if (page.cLevel > 0)
                    {
                        BTENTRY e;
                        unsafe
                        {
                            e = FM.TypeFromBuffer<BTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                        }
                        ReadBTtree(e.BREF.ib);
                    }
                    else
                    {
                        if (page.pageTrailer.ptype == Eptype.ptypeNBT)
                        {
                            unsafe
                            {
                                NBTENTRY nbtEntry = FM.TypeFromBuffer<NBTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                                NBTs.Add(CleanEntry(nbtEntry));
                            }
                        }
                        else if (page.pageTrailer.ptype == Eptype.ptypeBBT)
                        {
                            unsafe
                            {
                                BBTENTRY bbtEntry = FM.TypeFromBuffer<BBTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                                BBTs.Add(CleanEntry(bbtEntry));
                            }
                        }
                        else
                            throw new Exception("Unexpected page entry type");
                    }
                }
            }
            catch (Exception ex)
            {
                stream?.Close();
                throw new Exception($"ERROR READING NBT TREE: {ex.Message}");
            }
        }
        private void ReadBTtreeOST(ulong offset)
        {
            try
            {
                var page = FM.ReadType<OST_BTPAGE>(stream, offset);
                for (int i = 0; i < page.cEnt; i++)
                {
                    if (page.cLevel > 0)
                    {
                        BTENTRY e;
                        unsafe
                        {
                            e = FM.TypeFromBuffer<BTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                        }
                        ReadBTtreeOST(e.BREF.ib);
                    }
                    else
                    {
                        if (page.pageTrailer.ptype == Eptype.ptypeNBT)
                        {
                            unsafe
                            {
                                NBTENTRY nbtEntry = FM.TypeFromBuffer<NBTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                                NBTs.Add(CleanEntry(nbtEntry));
                            }
                        }
                        else if (page.pageTrailer.ptype == Eptype.ptypeBBT)
                        {
                            unsafe
                            {
                                OST_BBTENTRY ostEntry = FM.TypeFromBuffer<OST_BBTENTRY>(page.rgentries, pageEntries, i * page.cbEnt);
                                BBT bbt = CleanEntry(ostEntry);
                                BBTs.Add(bbt);
                            }
                        }
                        else
                            throw new Exception("Unexpected page entry type");
                    }
                }
            }
            catch (Exception ex)
            {
                stream?.Close();
                throw new Exception($"ERROR READING NBT TREE: {ex.Message}");
            }
        }
        private NBTENTRY CleanEntry(NBTENTRY entry)
        {
            NBTENTRY nbt = new NBTENTRY()
            {
                nid = entry.nid,
                dwPad = 0,
                bidData = PST.cleanBid(entry.bidData),
                bidSub = PST.cleanBid(entry.bidSub),
                nidParent = entry.nidParent,
                dwPadding = 0
            };
            return nbt;
        }
        private BBT CleanEntry(BBTENTRY entry)
        {
            entry.BREF.bid = PST.cleanBid(entry.BREF.bid);
            int cLevel = 0; int cEnt = 0; UInt32 IcbTotal = 0;
            EbType bType = GetBlockType(entry, ref cLevel, ref cEnt, ref IcbTotal);
            BBT bbt = new BBT()
            {
                type = bType,
                BREF = entry.BREF,
                cbStored = entry.cb,
                cbInflated = entry.cb,
                dwPadding = 0,
                cRef = entry.cRef,
                cEnt = cEnt,
                cLevel = cLevel,
                IcbTotal = IcbTotal,
            };
            return bbt;
        }
        private BBT CleanEntry(OST_BBTENTRY entry)
        {
            entry.BREF.bid = PST.cleanBid(entry.BREF.bid);
            int cLevel = 0; int cEnt = 0; UInt32 IcbTotal = 0;
            EbType bType = GetBlockType(entry, ref cLevel, ref cEnt, ref IcbTotal);
            BBT bbt = new BBT()
            {
                type = bType,
                BREF = entry.BREF,
                cbStored = entry.cbStored,
                cbInflated = entry.cbInflated,
                dwPadding = 0,
                cRef = entry.cRef,
                cEnt = cEnt,
                cLevel = cLevel,
                IcbTotal = IcbTotal,
            };
            return bbt;
        }
        public List<byte[]> ReadDatablocks(UInt64 bid)
        {   
            List<byte[]> datablocks = new List<byte[]>();
            int index = BBTs.FindIndex(x => x.BREF.bid == bid);
            if (IsExternal(bid))
            {
                datablocks.Add(ReadDatablock(BBTs[index]));
            }
            else
            {
                foreach (byte[] block in GetNextDatatreeBlock(BBTs[index]))
                {
                    datablocks.Add(block);
                }
            }
            return datablocks;
        }
        private IEnumerable<byte[]> GetNextDatatreeBlock(BBT bbt)
        {
            foreach (byte[] newBlock in GetNextBlock(bbt.BREF.bid))
                yield return newBlock;
            yield break;
        }
        private IEnumerable<byte[]> GetNextBlock(UInt64 bid)
        {
             // internal datablock - X or XL
             byte[] block = ReadDatablock(bid);
             byte bType = FM.ExtractTypeFromArray<byte>(block, 0);
             byte cLevel = FM.ExtractTypeFromArray<byte>(block, 1);
             UInt16 cEnt = FM.ExtractTypeFromArray<UInt16>(block, 2);
            if (cLevel == 1)
            {
                for (int i = 0; i < cEnt; i++)
                {   // get block
                    UInt64 dBid = PST.cleanBid(FM.ExtractTypeFromArray<UInt64>(block, 8 + 8 * i));
                    yield return ReadDatablock(dBid);
                }
            }
            else
            {  // xxblock level 2
                for (int i = 0; i < cEnt; i++)
                {
                    UInt64 xBid = PST.cleanBid(FM.ExtractTypeFromArray<UInt64>(block, 8 + 8 * i));
                    foreach (byte[] newBlock in GetNextBlock(xBid))
                        yield return newBlock;
                }
            }
            yield break;
        }
        public byte[] ReadFullDatablock(UInt64 bid)
        {
            if ((bid & 0x02) != 0)
            {
                return FM.srcFile.ReadDatatree(bid);
            }
            else
            {
               return FM.srcFile.ReadDatablock(bid);
            }
        }
        public byte[] ReadDatablock(UInt64 bid)
        {
            int index = BBTs.FindIndex(x => x.BREF.bid == bid);
            return ReadDatablock(BBTs[index]);
        }
        public byte[] ReadDatablock(BBTENTRY bbtEntry)
        {
            byte[] datablock = new byte[bbtEntry.cb];
            stream.Position = (long)bbtEntry.BREF.ib;
            int bytesRead = stream.Read(datablock, 0, bbtEntry.cb);
            if (bytesRead != bbtEntry.cb) { throw new Exception($"DATABLOCK READ ERROR (bid {bbtEntry.BREF.bid})"); }
            if (header.bCryptMethod == EbCryptMethod.NDB_CRYPT_PERMUTE && IsExternal(bbtEntry))
            {
                PST.CryptPermuteDecode(datablock);
            }
            return datablock;
        }
        public byte[] ReadDatablock(OST_BBTENTRY bbtEntry)
        {
            byte[] datablock;
            stream.Position = (long)bbtEntry.BREF.ib;
            if (bbtEntry.cbStored != bbtEntry.cbInflated)
            {
                // The first two bytes are a zlib header which DeflateStream does not understand
                // They should be 0x789c, the magic code for default compression
                if (stream.ReadByte() != 0x78 || stream.ReadByte() != 0x9c)
                    throw new Exception("OST DATA BLOCK ERROR: Unexpected header in compressed data stream");
                using (DeflateStream decompressionStream = new DeflateStream(stream, CompressionMode.Decompress, true))
                {
                    datablock = new byte[bbtEntry.cbInflated];
                    decompressionStream.Read(datablock, 0, bbtEntry.cbInflated);
                }
            }
            else
            {
                datablock = new byte[bbtEntry.cbStored];
                stream.Read(datablock, 0, bbtEntry.cbStored);
            }
            if ((bbtEntry.BREF.bid & 0x02) == 0 & header.bCryptMethod == EbCryptMethod.NDB_CRYPT_PERMUTE)
            {
                // Key for cyclic algorithm is the low 32 bits of the BID, so supply it in case it's needed
                PST.CryptPermuteDecode(datablock);
            }
            return datablock;
        }
        public byte[] ReadDatablock(BBT bbtEntry)
        {
            byte[] datablock;
            stream.Seek((long)bbtEntry.BREF.ib, SeekOrigin.Begin);
            int offset = 0;
            if (bbtEntry.cbStored != bbtEntry.cbInflated)
            {
                // The first two bytes are a zlib header which DeflateStream does not understand
                // They should be 0x789c, the magic code for default compression
                if (stream.ReadByte() != 0x78 || stream.ReadByte() != 0x9c)
                    throw new Exception("OST DATA BLOCK ERROR: Unexpected header in compressed data stream");
                using (DeflateStream decompressionStream = new DeflateStream(stream, CompressionMode.Decompress, true))
                {
                    int totalRead = 0;
                    int toBeRead = bbtEntry.cbInflated;
                    datablock = new byte[bbtEntry.cbInflated];
                    while (totalRead < bbtEntry.cbInflated)
                    {
                        int bytesRead = decompressionStream.Read(datablock, totalRead, toBeRead);
                        if (bytesRead == 0) break;
                        totalRead += bytesRead;
                        toBeRead -= bytesRead;
                    }
                }
                // decompression stream no longer reads all requested bytes on donet 6+
                // https://learn.microsoft.com/en-us/dotnet/core/compatibility/core-libraries/6.0/partial-byte-reads-in-streams
            }
            else
            {
                datablock = new byte[bbtEntry.cbStored];
                stream.Read(datablock, 0, bbtEntry.cbStored);
            }
            if ((bbtEntry.BREF.bid & 0x02) == 0 & header.bCryptMethod == EbCryptMethod.NDB_CRYPT_PERMUTE)
            {
                // Key for cyclic algorithm is the low 32 bits of the BID, so supply it in case it's needed
                PST.CryptPermuteDecode(datablock);
            }
            return datablock;
        }
        public byte[] ReadDatatree(UInt64 bid)
        {
            if (IsExternal(bid)) throw new Exception($"ENTRY ERROR...(bid {bid}) DATATREE SHOULD BE internal datablock");
            byte[] datablock = new byte[0];
            ReadDatatree(bid, ref datablock);
            return datablock;
        }
        public void ReadDatatree(UInt64 dataBid, ref byte[] datablock)
        {
            byte[] newBlock = ReadDatablock(dataBid);
            if (IsExternal(dataBid))
            {
                int dSize = datablock.Length;
                Array.Resize(ref datablock, dSize + newBlock.Length);
                newBlock.CopyTo(datablock, dSize);
            }
            else
            {  // internal datablock - X or XL
                byte bType = FM.ExtractTypeFromArray<byte>(newBlock, 0);
                byte cLevel = FM.ExtractTypeFromArray<byte>(newBlock, 1);
                UInt16 cEnt = FM.ExtractTypeFromArray<UInt16>(newBlock, 2);
                for (int i = 0; i < cEnt; i++)
                {
                    UInt64 bid = PST.cleanBid(FM.ExtractTypeFromArray<UInt64>(newBlock, 8 + 8 * i));
                    ReadDatatree(bid, ref datablock);
                }
            }
        }
        public List<SLENTRY> GetSLentries(UInt64 bid)
        {
            List<SLENTRY> slEntries = new List<SLENTRY>();
            if (IsExternal(bid)) throw new Exception($"ENTRY ERROR...(bid {bid}) DATATREE SHOULD BE internal datablock");
            GetSLentries(bid, ref slEntries);
            return slEntries;
        }
        private void GetSLentries(UInt64 bid, ref List<SLENTRY> slEntries)
        {
            if (bid == 0) return;
            byte[] sBlock = ReadDatablock(bid);
            byte bType = FM.ExtractTypeFromArray<byte>(sBlock, 0);
            byte cLevel = FM.ExtractTypeFromArray<byte>(sBlock, 1);
            UInt16 cEnt = FM.ExtractTypeFromArray<UInt16>(sBlock, 2);
            if (bType != (byte)EbType.bTypeS) throw new Exception($"ENTRY ERROR...(bid {bid}) IS NOT a SL tree");
            if (cLevel == 0)
            {
                for (int i = 0; i < cEnt; i++)
                {
                    SLENTRY slEntry = FM.ExtractTypeFromArray<SLENTRY>(sBlock, 8 + 24 * i);
                    slEntry.nid = PST.cleanNID(slEntry.nid);
                    slEntry.bidData = PST.cleanBid(slEntry.bidData);
                    slEntry.bidSub = PST.cleanBid(slEntry.bidSub);
                    if (slEntry.nid != 0)
                    {
                        if (BBTs.FindIndex(x => x.BREF.bid == slEntry.bidData) < 0)
                        {
                            Program.mainForm.statusMSG($"SL block (bid={bid}) with index={i}, nid={slEntry.nid}:  bidData={slEntry.bidData} not defined. Set to zero");
                            slEntry.bidData = 0;
                        }
                        if (slEntry.bidSub > 0 & BBTs.FindIndex(x => x.BREF.bid == slEntry.bidSub) < 0)
                        {
                            Program.mainForm.statusMSG($"SL block (bidSub={bid}) with index={i}, nid={slEntry.nid}:  bidSub={slEntry.bidSub} not defined. Set to zero");
                            slEntry.bidSub = 0;
                        }
                        slEntries.Add(slEntry);
                    }
                }
            }
            else
            {  // must be level 1 = SI block!!!
                for (int i = 0; i < cEnt; i++)
                {
                    SIENTRY siEntry = FM.ExtractTypeFromArray<SIENTRY>(sBlock, 8 + 16 * i);
                    siEntry.bid = PST.cleanBid(siEntry.bid);
                    GetSLentries(siEntry.bid, ref slEntries);
                }
            }
        }
        private bool IsExternal(UInt64 bid) { return (bid & 0x02) == 0; }
        private bool IsExternal(BBTENTRY bbtEntry) { return (bbtEntry.BREF.bid & 0x02) == 0; }
        private EbType GetBlockType(OST_BBTENTRY bbt, ref int cLevel, ref int cEnt, ref UInt32 IcbTotal)
        {
            EbType type = EbType.bTypeD; // external data
            if (bbt.BREF.bid % 4 != 0)
            {  // internal datablock x/xx/sl/si
                byte[] datablock = ReadDatablock(bbt);
                type = (EbType)datablock[0];
                cLevel = (int)datablock[1];
                cEnt = FM.ExtractTypeFromArray<UInt16>(datablock, 2);
                if (type == EbType.bTypeX)
                {
                    IcbTotal = FM.ExtractTypeFromArray<UInt32>(datablock, 4);
                }
            }
            return type;
        }
        private EbType GetBlockType(BBTENTRY bbt, ref int cLevel, ref int cEnt, ref UInt32 IcbTotal)
        {
            EbType type = EbType.bTypeD; // external data
            if (bbt.BREF.bid % 4 != 0)
            {  // internal datablock x/xx/sl/si
                byte[] datablock = ReadDatablock(bbt);
                type = (EbType)datablock[0];
                cLevel = (int)datablock[1];
                cEnt = FM.ExtractTypeFromArray<UInt16>(datablock, 2);
                if (type == EbType.bTypeX)
                {
                    IcbTotal = FM.ExtractTypeFromArray<UInt32>(datablock, 4);
                }
            }
            return type;
        }
    }
}
