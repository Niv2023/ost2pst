// BTree manager for the PST writer: for NBT and BBT trees
// 

namespace ost2pst
{
    //  Generic handling for all trees
    public class NTree
    {
        private Eptype type;
        private PstFile pst;
        private int maxNrOfLevelsPerTree;
        private int maxNrBranchesPerNode;
        private int maxNrLeavesPerNode;
        private int nrOfLeaves;
        private List<BTENTRY> branchNodes;
        private List<BTENTRY> upperBranchNodes;
        private List<NBTENTRY> NBTs;
        private List<BBTENTRY> BBTs;
        private const int cEntMaxPercent = 100;  // 100% allocation on node entry pages - can be set to a lower value in the future
        private int nrOfLeafNodes;
        public NTree(List<NBTENTRY> nbts)
        {
            type = Eptype.ptypeNBT;
            initNodes(MS.NBTcbEnt);
            NBTs = nbts;
            nrOfLeaves = NBTs.Count;
        }
        public NTree(List<BBTENTRY> bbts)
        {
            type = Eptype.ptypeBBT;
            initNodes(MS.BBTcbEnt);
            BBTs = bbts;
            nrOfLeaves = BBTs.Count;
        }
        private void initNodes(int cbEnt)
        {
            float maxLimit = 1; //100%
            if (cEntMaxPercent > 0 && cEntMaxPercent < 100) { maxLimit = (float)cEntMaxPercent / 100; }
            maxNrOfLevelsPerTree = MS.MaxNrTreeLevels;
            maxNrBranchesPerNode = (int)((MS.BTPageEntryBytes / MS.BBTcbEnt) * maxLimit);
            maxNrLeavesPerNode = MS.BTPageEntryBytes / cbEnt;
            maxNrLeavesPerNode = (int)(maxNrLeavesPerNode * maxLimit);
            branchNodes = new List<BTENTRY>();
            upperBranchNodes = new List<BTENTRY>();
            nrOfLeafNodes = nrOfLeaves / maxNrLeavesPerNode;
            if (nrOfLeaves % maxNrLeavesPerNode > 0) nrOfLeafNodes++;
        }

        public BREF ExportNodes(PstFile pstFile) {
            pst = pstFile;
            if (type == Eptype.ptypeBBT) { exportBBTLeafNodes(); } else { exportNBTLeafNodes(); }
            return exportBranchNodes();
        }
        private BREF exportBranchNodes()
        {
            return exportBranchLevel(1);
        }
        private BREF exportBranchLevel(int level)
        {
             if (branchNodes.Count == 1)
            {
                // top of branch tree
                return branchNodes[0].BREF;
            }
            else
            {
                List<BTENTRY> branchPageEntries = new List<BTENTRY>();
                for (int i = 0; i < branchNodes.Count; i++)
                {
                    branchPageEntries.Add(branchNodes[i]);
                    if (branchPageEntries.Count % maxNrBranchesPerNode == 0)
                    {
                        BTENTRY btEntry = new BTENTRY()
                        {
                            btkey = branchPageEntries[0].btkey,
                            BREF = pst.AddPage(branchPageEntries, type, level)
                        };
                        upperBranchNodes.Add(btEntry);
                        branchPageEntries = new List<BTENTRY>();
                    }
                }
                if (branchPageEntries.Count > 0)
                {
                    BTENTRY btEntry = new BTENTRY()
                    {
                        btkey = branchPageEntries[0].btkey,
                        BREF = pst.AddPage(branchPageEntries, type, level)
                    };
                    upperBranchNodes.Add(btEntry);
                }
                branchNodes = upperBranchNodes;
                upperBranchNodes = new List<BTENTRY>();
                return exportBranchLevel(level + 1);
            }
        }
        private void exportBBTLeafNodes()
        {
            int nb = 1;
            List<BBTENTRY> leafPageEntries = new List<BBTENTRY>();
            for (int i = 0; i < nrOfLeaves; i++)
            {
                leafPageEntries.Add(BBTs[i]);
                if (leafPageEntries.Count % maxNrLeavesPerNode == 0)
                {
                    BTENTRY btEntry = new BTENTRY()
                    {
                        btkey = leafPageEntries[0].BREF.bid,
                        BREF = pst.AddPage(leafPageEntries)
                    };
                    branchNodes.Add(btEntry);
                    leafPageEntries = new List<BBTENTRY>();
                }
            }
            if (leafPageEntries.Count > 0)
            {
                BTENTRY btEntry = new BTENTRY()
                {
                    btkey = leafPageEntries[0].BREF.bid,
                    BREF = pst.AddPage(leafPageEntries)
                };
                branchNodes.Add(btEntry);
            }
         }
        private void exportNBTLeafNodes()
        {
            int nb = 1;
            List<NBTENTRY> leafPageEntries = new List<NBTENTRY>();
            for (int i = 0; i < nrOfLeaves; i++)
            {
                leafPageEntries.Add(NBTs[i]);
                if (leafPageEntries.Count % maxNrLeavesPerNode == 0)
                {
                    BTENTRY btEntry = new BTENTRY()
                    {
                        btkey = leafPageEntries[0].nid.dwValue,
                        BREF = pst.AddPage(leafPageEntries)
                    };
                    branchNodes.Add(btEntry);
                    leafPageEntries = new List<NBTENTRY>();
                }
            }
            if (leafPageEntries.Count > 0)
            {
                BTENTRY btEntry = new BTENTRY()
                {
                    btkey = leafPageEntries[0].nid.dwValue,
                    BREF = pst.AddPage(leafPageEntries)
                };
                branchNodes.Add(btEntry);
            }
        }
    }
}