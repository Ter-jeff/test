using System.Linq;

namespace Automation.Reader.ScghFile.ProCharPatternSet.Base
{

    public class ProdCharPatternSetHardIp : ProdCharPatternSetBase
    {
        public ProdCharPatternSetHardIp(ProdCharPatternSetBase basePatset)
        {
            InitList = basePatset.InitList;
            InitAliasList = basePatset.InitAliasList.ToList();
            PayloadList = basePatset.PayloadList.ToList();
            PayloadAliasList = basePatset.PayloadAliasList.ToList();
            InstanceName = basePatset.InstanceName;
            PatSetName = basePatset.PatSetName;
            InitPatSetNameByNamingRule = basePatset.InitPatSetNameByNamingRule;
            InitPatternMissing = basePatset.InitPatternMissing;
            ProdCharRow = basePatset.ProdCharRow;
        }
    }
}
