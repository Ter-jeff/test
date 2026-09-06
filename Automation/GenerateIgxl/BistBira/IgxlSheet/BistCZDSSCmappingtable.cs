using System.Collections.Generic;
using System.Linq;

namespace Automation.GenerateIgxl.BistBira.IgxlSheet
{
    public class BistCzDsscMappingTable
    {
        public List<BistCzDsscOnePattern> BistCzDsscOnePatterns = new List<BistCzDsscOnePattern>();
        public string OutputFile;
    }

    public class BistCzDsscOnePattern
    {
        public string Block;
        public List<BistCzDsscCategory> BistCzDsscCategories = new List<BistCzDsscCategory>();
    }

    public class BistCzDsscCategory
    {
        public string CategoryName;
        public List<BistCzDsscCategoryRow> BistCzDsscCategoryRows = new List<BistCzDsscCategoryRow>();

        public BistCzDsscCategory() { }

        public BistCzDsscCategory(BistCzDsscCategory other)
        {
            if (other == null)
            {
                return;
            }

            CategoryName = other.CategoryName;

            BistCzDsscCategoryRows = other.BistCzDsscCategoryRows?.Select(row => row.Copy()).ToList() ?? new List<BistCzDsscCategoryRow>();
        }

        public BistCzDsscCategory Copy()
        {
            return new BistCzDsscCategory(this);
        }
    }

    public class BistCzDsscCategoryRow
    {
        public string CategoryName;
        public string PatternName;
        public string CycleNumber;
        public string SegmentNumber;
        public string Description;
        public string TestAutoSwitch;
        public string Values;

        public BistCzDsscCategoryRow() { }

        public BistCzDsscCategoryRow(BistCzDsscCategoryRow other)
        {
            if (other == null)
            {
                return;
            }

            CategoryName = other.CategoryName;
            PatternName = other.PatternName;
            CycleNumber = other.CycleNumber;
            SegmentNumber = other.SegmentNumber;
            Description = other.Description;
            TestAutoSwitch = other.TestAutoSwitch;
            Values = other.Values;
        }

        public BistCzDsscCategoryRow Copy()
        {
            return new BistCzDsscCategoryRow(this);
        }
    }
}
