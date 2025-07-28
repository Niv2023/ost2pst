using System.Data;

namespace ost2pst
{
    //
    // Create BTH tree

    //    public class BTH : Node
    public class BTH
    {
        private BTHHEADER header;
        public int maxBthLeaves;
        public int maxBthBranches;
        public int cbKey;
        public int cbEnt;
        public UInt16 hidIx;
        public HID[] hnHIDs;
        List<Node> nodes;
        private int hnBlockSize;
        public UInt16 hidBlockIx;
        public UInt32 subnodeNID;
        public UInt32 rowMatrixSubnode;
        private bool newSubnodes;
        public List<SLENTRY> subnodes = new List<SLENTRY>();
        public Branch root;
        private TableContext tc;
        private List<Property> props;
        private List<TCROWID> rowids;
        public List<HeapItem> heapItems;
        public List<Property> propsInTCheap;
        public int lastKey = 0;
        private List<RowCellRef> hidCells;

        // Constructor for PC BTH's
        public BTH(List<Property> properties)
        {
            props = properties;
            root = new Branch(this, 0, 0);
            cbKey = 2;
            cbEnt = 6;
            header = new BTHHEADER()
            {
                btype = EbType.bTypeBTH,
                cbKey = 2,
                cbEnt = 6,
                hidRoot = new HID(EnidType.HID, 2, 0)
            };
            BuiltPCBTHtree();
            header.bIdxLevels = (byte)root.level;
            GetPCHeapNodes();
            GetHeapItems();
            subnodes = subnodes.OrderBy(s => s.nid).ToList();
        }
        private unsafe Node BTHheaderNode()
        {
            byte[] bh = new byte[sizeof(BTHHEADER)];
            FM.TypeToArray<BTHHEADER>(ref bh, header);
            return new Hdata(bh, 1);  // BTH header first on the heap for PC and TC
        }
        private unsafe Node TCheaderNode()
        {
            byte[] th = tc.Header();
            return new Hdata(th, 2);  // TC heard on the 2nd position
        }
        // Constructor for PC BTH's
        public BTH(TableContext TC, bool withSubnodes = false)
        {
            tc = TC;
            newSubnodes = withSubnodes;
            cbKey = 4;
            cbEnt = 4;
            if (newSubnodes)
            {
                subnodeNID = 1;
                if (tc.tcINFO.hnidRows.IsNID)
                {
                    rowMatrixSubnode = tc.tcINFO.hnidRows.dwValue;
                    subnodeNID = rowMatrixSubnode >> 5;
                }
            }
            root = new Branch(this, 0, 0);
            header = new BTHHEADER()
            {
                btype = EbType.bTypeBTH,
                cbKey = 4,
                cbEnt = 4,
                hidRoot = new HID(EnidType.HID, 0, 0),
                bIdxLevels = 0
            };
            if (tc.tcRowIndexes.Count == 0)
            {  // TC without row data
                nodes = [BTHheaderNode(), TCheaderNode()];
            }
            else
            {
                BuiltTCBTHtree();
                GetTCHeapNodes();
            }
            GetHeapItems();
            subnodes = subnodes.OrderBy(s => s.nid).ToList();
        }
        private void BuiltTCBTHtree()
        {
            root = new Branch(this, 0, 0);
            hidCells = new List<RowCellRef>();
            maxBthLeaves = MS.MaxHeapItemSize / (header.cbKey + header.cbEnt);
            maxBthBranches = MS.MaxHeapItemSize / (header.cbKey + 4);   // sizeof(HID) = 4
            header.hidRoot = new HID(EnidType.HID, 3, 0);    // hid=0x60 for the RowIndex
            foreach (TCROWID r in tc.tcRowIndexes)
            {
                Leaf tcLeaf = new Leaf(this, r);
                if (!root.AddLeaf(tcLeaf))
                { // root branch full create a new level
                    Branch newRoot = new Branch(this, root.key, root.level + 1);
                    newRoot.AddBranch(root);
                    newRoot.AddLeaf(tcLeaf);
                    root = newRoot;
                }
            }
            header.bIdxLevels = (byte)root.level;
        }
        private void BuiltPCBTHtree()
        {   // builds the pc BTH tree adding subnodes if required
            maxBthLeaves = MS.MaxHeapItemSize / (header.cbKey + header.cbEnt);
            maxBthBranches = MS.MaxHeapItemSize / (header.cbKey + 4);   // sizeof(HID) = 4
            foreach (Property p in props)
            {
                Leaf pcLeaf = new Leaf(this, p);
                if (!root.AddLeaf(pcLeaf))
                { // root branch full create a new level
                    Branch newRoot = new Branch(this, root.key, root.level + 1);
                    newRoot.AddBranch(root);
                    newRoot.AddLeaf(pcLeaf);
                    root = newRoot;
                }
            }
        }
        private void GetPCHeapNodes()
        {
            nodes = [BTHheaderNode()];
            AddBTHnodeItems();
            AddPCBTHdataItems(root);
            UpdadeHNIDs();
            UpdatePCBTHhnids(root);
        }
        private void GetTCHeapNodes()
        {
            nodes = [BTHheaderNode(), TCheaderNode()];
            AddBTHnodeItems();
            AddTCBTHdataItems();
            if (tc.RowMatrixSize <= MS.MaxHeapItemSize)
            {   // rowdata will be added in the heap
                // add blank array now just for calculating the hid's
                Hdata rowM = new Hdata(new byte[tc.RowMatrixSize]);
                AddNode(rowM);
            }
            UpdadeHNIDs();
            UpdateTCrowdataHnids();
            if (tc.RowMatrixSize <= MS.MaxHeapItemSize)
            {  // add row data in the heap
                nodes[^1].data = tc.RowMatrixArray().First();
                tc.tcINFO.hnidRows = new HNID(nodes[^1].hid);
            }
            else
            {  // rowdata in a subnode
                if (newSubnodes)
                {
                    NID sn = new NID(rowMatrixSubnode);
                    subnodes.Add(AddSubnode (sn,tc.RowMatrixArray()));
                }
            }
            nodes[1].data = TCheaderNode().data; // overwrite with the updated header
        }
        private void AddBTHnodeItems()
        {
            AddNode(root);
            // add nodes top down the tree
            for (int level = root.level - 1; level >= 0; level--)
            {
                AddNodesOnLevel(level, root);
            }
        }
        private void AddNode(Node node)
        {
            node.hnIx = nodes.Count + 1;
            nodes.Add(node);
        }
        private void AddNodesOnLevel(int level, Branch node)
        {
            foreach (Node n in node.children)
            {
                if (n.level == level)
                {
                    AddNode(n);
                }
                else
                {
                    AddNodesOnLevel(level, (Branch)n);
                }
            }
        }
        private void AddPCBTHdataItems(Branch node)
        {
            if (node.level == 0)
            {
                foreach (Node n in node.children)
                {
                    // PCBTH leaf... check if data required on heap
                    Leaf l = (Leaf)n;
                    if (l.pcbth.dwValueHnid.HasValue)
                    {
                        if (l.hid.hidType == EnidType.HID)
                        {   // add a heap node for the property data
                            AddNode(new Hdata(l));
                            l.hnIx = nodes.Count;
                        }
                    }
                }
            }
            else
            {
                foreach (Node n in node.children)
                {
                    AddPCBTHdataItems((Branch)n);
                }
            }
        }
        private void UpdatePCBTHhnids(Branch node)
        {
            if (node.level == 0)
            {
                foreach (Node n in node.children)
                {
                    // PCBTH leaf... 
                    Leaf l = (Leaf)n;
                    if (l.hid.hidType == EnidType.HID)
                    {
                        l.hid = hnHIDs[l.hnIx];
                    }
                }
            }
            else
            {
                foreach (Node n in node.children)
                {
                    UpdatePCBTHhnids((Branch)n);
                }
            }
        }
        private void UpdateTCrowdataHnids()
        {
            foreach (RowCellRef cell in hidCells)
            {
                RowData rd = tc.tcRowMatrix[cell.rowIndex];
                HID hid = hnHIDs[cell.nodeIx];
                FM.TypeToArray<UInt32>(ref rd.rgbData, hid.hidValue, cell.rowPos);
            }
        }
        private void AddTCBTHdataItems()
        {   // TC heap items are refered on the data rows
            propsInTCheap = new List<Property>();
            for (int i = 0; i < tc.tcRowMatrix.Count; i++)
            {
                RowData rd = tc.tcRowMatrix[i];
                for (int j = 0; j < rd.Props.Count; j++)
                {
                    Property prop = rd.Props[j];
                    int col = tc.FindColumn(prop.id, prop.type);
                    if (col < 0)
                    {   //// *************** tbd
                        continue;
                        //throw new Exception($"Property id {prop.id} not in the TC coldesc");
                    }
                    int cb = tc.tcCols[col].cbData;
                    int ib = tc.tcCols[col].ibData;
                    int bit = tc.tcCols[col].iBit;
                    if (prop.hasFixedSize & prop.dataSize <= 8)
                    {  // fits in the data row
                        PST.SetBitInArray(ref rd.rgbCEB, bit);
                        if (prop.id != EpropertyId.PidTagLtpRowVer)
                        {   // don't copy ost row verstion but use the new range for pst
                            Array.Copy(prop.data, 0, rd.rgbData, ib, prop.data.Length);
                        }
                    }
                    else
                    {
                        if (prop.hnid.HasValue  || prop.dataSize > 0)
                        {
                            PST.SetBitInArray(ref rd.rgbCEB, bit);
                            if (cb != 4) throw new Exception($"Property id {prop.id} cb={cb} expected 4");
                            if (prop.dataSize > MS.MaxHeapItemSize)
                            {  // data should be placed in a subnode
                                if (newSubnodes)
                                {
                                    subnodeNID++;
                                    NID sn = new NID(EnidType.LTP, subnodeNID);
                                    FM.TypeToArray<UInt32>(ref rd.rgbData, sn.dwValue, ib);
                                    subnodes.Add(AddSubnode(sn, prop.data));
                                }
                                else
                                {  // use the original nid
                                    FM.TypeToArray<UInt32>(ref rd.rgbData, prop.hnid.dwValue, ib);
                                }
                            }
                            else
                            {  // place in the heap
                                if (prop.dataSize > 0)
                                {
                                    AddNode(new Hdata(prop.data));
                                    hidCells.Add(new RowCellRef() { rowIndex = i, rowPos = ib, nodeIx = nodes.Count });
                                }
                            }
                        }
                        else
                        {   // scanpst flags   !!Receive folder table missing default message class
                            // when the rgbCEB is not set... outlook seems to accept it
                            if (prop.id == EpropertyId.PidTagMessageClassW) PST.SetBitInArray(ref rd.rgbCEB, bit);
                        }
                    }
                }

            }
        }

        private void UpdadeHNIDs()
        {
            CheckHIDsBlockIndexes();
            UpdateNodeHIDs();
        }
        private void CheckHIDsBlockIndexes()
        {
            hnHIDs = new HID[nodes.Count + 1];  // HID are 1 base
            hidBlockIx = 0;
            hidIx = 0;
            hnBlockSize = PST.HNdatablockUsage(hidBlockIx);
            foreach (Node n in nodes)
            {
                int nodeSize = n.Size + 2;  // each entry requires 2 bytes in the rgibAlloc
                if (hnBlockSize + nodeSize > MS.MaxDatablockSize)
                {   // allocate on a new data block
                    hidIx = 0;
                    hidBlockIx++;
                    hnBlockSize = PST.HNdatablockUsage(hidBlockIx);
                }
                hnBlockSize += nodeSize;
                hidIx++;
                hnHIDs[n.hnIx] = new HID(EnidType.HID, hidIx, hidBlockIx);
            }
        }
        private void UpdateNodeHIDs()
        {
            foreach (Node n in nodes)
            {
                n.hid = hnHIDs[n.hnIx];
            }
        }

        private unsafe void GetHeapItems()
        {
            heapItems = new List<HeapItem>();
            for (int i = 0; i < nodes.Count; i++)
            {
                heapItems.Add(new HeapItem(nodes[i].hid, nodes[i].Data));
            }
        }
        public SLENTRY AddSubnode(NID sNid, byte[] data)
        {
            BREF bref = Blocks.AddDatablock(data);
            return new SLENTRY()
            {
                nid = sNid.dwValue,
                bidData = bref.bid,
                bidSub = 0
            };
        }
        public SLENTRY AddSubnode(NID sNid, List<byte[]> data)
        {
            if(data.Count == 1)
            {
                return AddSubnode(sNid, data[0]);
            }
            else
            {
                BREF bref = Blocks.AddDataTree(data);
                return  new SLENTRY()
                {
                    nid = sNid.dwValue,
                    bidData = bref.bid,
                    bidSub = 0
                };
            }
        }
    }
    #region node classes
    // each bth element (Leave or intermidiate node) has a HID
    // reference to the BTH tree class for helper functions
   
    public class RowCellRef
    {
        public int rowIndex;
        public int rowPos;
        public int nodeIx;
    }
    public abstract class Node
    {
        public BTH bth;  // reference to the BTH object
        public int level;
        public UInt32 key;
        public int hnIx; // allocation order in the heapnode (1 based)
        public HID hid;
        public byte[] data;
        public abstract int Size { get; }
        public abstract byte[] Data { get; set; }
    }

    // data on heap node

    // BTH leaves for PC has the PCBTH
    public class Hdata : Node
    {
        public unsafe override int Size
        {
            get
            {   // key+hid or key+data 8 bytes (2+6 or 4+4)
                return data.Length;
            }
        }
        public unsafe override byte[] Data
        {
            get
            {   // key+hid or key+data 8 bytes (2+6 or 4+4)
                return data;
            }
            set
            {
                data = value;
            }
        }
        public Hdata(Leaf l, int hnx)
        {
            data = l.property.data;
            level = -1;
            hid = l.hid;
            hnIx = hnx;
        }
        public Hdata(Leaf l)
        {
            data = l.property.data;
            level = -1;
            hid = l.hid;
            hnIx = 0;
        }
        public Hdata(byte[] bytes)
        {
            data = bytes;
            level = -1;
            hid = new HID(EnidType.HID, 0, 0);
            hnIx = 0;
        }
        public Hdata(byte[] bytes, int hnx)
        {
            data = bytes;
            level = -1;
            hid = new HID(EnidType.HID, (UInt16)hnx, 0);
            hnIx = hnx;
        }
        public Hdata(byte[] bytes, HID Hid)
        {
            data = bytes;
            level = -1;
            hid = Hid;
            hnIx = hid.GetIndex(false);
        }
    }
    // BTH leaves for PC has the PCBTH
    public class Leaf : Node
    {
        public PCBTH pcbth;
        public TCROWID tcrowid;
        public Property property;
        public unsafe override int Size
        {
            get
            {   // key+hid or key+data 8 bytes (2+6 or 4+4)
                return sizeof(PCBTH);
            }
        }
        public unsafe override byte[] Data
        {
            get
            {   // key+hid or key+data 8 bytes (PC=2+6 or TC=4+4)
                data = new byte[8];  // pcbth record
                if (pcbth.wPropId != 0)
                {   // leaf is a PC
                    FM.TypeToArray<UInt16>(ref data, (UInt16)pcbth.wPropId, 0);
                    FM.TypeToArray<UInt16>(ref data, (UInt16)pcbth.wPropType, 2);
                    if (hid.hidType == EnidType.HID)
                    {
                        FM.TypeToArray<UInt32>(ref data, (UInt32)hid.hidValue, 4);
                    }
                    else
                    {
                        if (property.hasVariableSize && property.data.Length == 0)
                        {
                            FM.TypeToArray<UInt32>(ref data, 0, 4);
                        }
                        else
                        {
                            FM.TypeToArray<UInt32>(ref data, (UInt32)pcbth.dwValueHnid.dwValue, 4);
                        }
                    }
                }
                else
                {   // leaf is a row id
                    FM.TypeToArray<UInt32>(ref data, tcrowid.dwRowID, 0);
                    FM.TypeToArray<UInt32>(ref data, tcrowid.dwRowIndex, 4);
                }
                return data;
            }
            set
            {
                data = value;
            }
        }
        public Leaf(BTH Bth, Property prop)
        {
            bth = Bth;
            level = -1;
            key = (UInt16)prop.PCBTH.wPropId;
            bth.lastKey++;
            pcbth = prop.PCBTH;
            property = prop;
            if (prop.hasFixedSize & prop.dataSize <= 4)
            {  // hnid in pcbth already has the data value. Flag with a HID type internal
                hid = new HID(EnidType.INTERNAL, 0, 0);
            }
            else
            {   // variable size of fixed > 4... check hid or nid
                if (prop.dataSize > MS.MaxHeapItemSize || prop.PCBTH.dwValueHnid.IsNID)
                {  // data should be placed in a subnode
                    hid = new HID(prop.hnid);  // keep the same subnode nid
                }
                else
                {  // data to be placed in the heap - hid to be calculated
                    if (prop.data.Length != 0)
                    {
                        hid = new HID(EnidType.HID, 0, 0);
                    }
                    else
                    {
                        hid = new HID(EnidType.INTERNAL, 0, 0);
                    }
                }
            }
        }
        public Leaf(BTH Bth, TCROWID row)
        {
            bth = Bth;
            level = -1;
            key = row.dwRowID;
            bth.lastKey++;
            tcrowid = row;
        }
    }

    // The only thing that a tree node must have is a key (cbKey=4)
    public class Branch : Node
    {
        public List<Node> children = new List<Node>();
        public bool IsFull
        {
            get
            {
                if (level == 0)
                {
                    return children.Count >= bth.maxBthLeaves;
                }
                else
                {
                    return children.Count >= bth.maxBthBranches;
                }
            }
        }
        public bool IsNotFull => !IsFull;
        public override int Size
        {
            get
            {   // level 0 = 2 + 6 (PC) or 4+4 (TC)
                if (level == 0)
                {
                    return children.Count * (bth.cbKey + bth.cbEnt);
                }
                else
                {
                    return children.Count * (bth.cbKey + 4); // 4 = sizeof(HID)
                }
            }
        }
        public unsafe override byte[] Data
        {
            get
            {   // key+hid or key+data 8 bytes (2+4 or 2+6 or 4+4)
                byte[] data = new byte[Size];
                if (level == 0)
                {
                    for (int i = 0; i < children.Count; i++)
                    {
                        byte[] pcRec = children[i].Data;   // 8 bytes leaf data
                        Array.Copy(pcRec, 0, data, i * 8, pcRec.Length);
                    }
                }
                else
                {
                    for (int i = 0; i < children.Count; i++)
                    {
                        if (bth.cbKey == 2)
                        {  // pc
                            FM.TypeToArray<UInt16>(ref data, (UInt16)children[i].key, i * 6);
                            FM.TypeToArray<UInt32>(ref data, (UInt32)children[i].hid.hidValue, i * 6 + 2);
                        }
                        else
                        {  // tc
                            FM.TypeToArray<UInt32>(ref data, (UInt32)children[i].key, i * 8);
                            FM.TypeToArray<UInt32>(ref data, (UInt32)children[i].hid.hidValue, i * 8 + 4);
                        }
                    }
                }
                int s = level == 0 ? 8 : 6;
                return data;
            }
            set {; }
        }

        public Branch(BTH Bth, uint Key, int Level)
        {
            this.bth = Bth;
            key = Key;
            this.level = Level;
            children = new List<Node>();
        }
        public bool AddLeaf(Leaf leaf)
        {
            if (level == 0)
            {
                if (IsNotFull)
                {
                    if (children.Count == 0) { key = leaf.key; }   // branch key = to its first leaf child
                    children.Add(leaf);
                    //bth.BTHhnSize += 8;   // each leaf requires 8 bytes  
                    return true;
                }
                return false;
            }
            foreach (var child in children)
            {
                if (child is Branch branch && branch.AddLeaf(leaf))
                {
                    return true;
                }
            }
            if (IsNotFull)
            {
                Branch newBranch = new Branch(bth, key, level - 1);
                newBranch.key = leaf.key;
                newBranch.AddLeaf(leaf);
                children.Add(newBranch);
                return true;
            }
            return false;
        }
        public void AddBranch(Branch branch)
        {
            children.Add(branch);
        }
    }
    #endregion
}
