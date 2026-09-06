using BinCutScriptLib.Static;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.Base
{
    public class BinCutExpress
    {
        public string Title = "";
        public EnumPowerType PowerType
        {
            get
            {
                return BinCutConfig.PowerType[Title.ToUpper()];
            }
        }

        public bool IsEvaluatePin
        {
            get { return Express.Contains("EVALUATE", System.StringComparison.OrdinalIgnoreCase); }
        }
        public bool IsBinResult
        {
            get { return Express.Contains("BIN RESULT", System.StringComparison.OrdinalIgnoreCase); }
        }
        public bool IsProduct
        {
            get
            {
                return Express.Contains("PRODUCT", System.StringComparison.OrdinalIgnoreCase);
            }
        }
        public string Express = "";

        public EnumExpressType ExpressType
        {
            get
            {
                if (Reg.RegexLvCoreVoltage.IsMatch(Express)) //E1 Voltage
                {
                    return EnumExpressType.E1Voltage;
                }
                if (Reg.RegexLvCoreEvaluate.IsMatch(Express)) //Evaluation
                {
                    return EnumExpressType.Evaluate;
                }
                if (Reg.RegexLvCoreResult.IsMatch(Express))//Bin Result
                {
                    return EnumExpressType.BinResult;
                }
                if (Reg.RegexCoreHvcc.IsMatch(Express)) //HVCC
                {
                    return EnumExpressType.Hvcc;
                }
                if (Reg.RegexModeBinProductGb.IsMatch(Express)) //BinX Product - GB
                {
                    return EnumExpressType.BinProductGb;
                }
                if (Reg.RegexModeBinProduct.IsMatch(Express)) //Bin Product
                {
                    return EnumExpressType.BinProduct;
                }
                if (Reg.RegexCoreE1ProductGb.IsMatch(Express)) //E1 Product
                {
                    return EnumExpressType.E1ProductGb;
                }
                if (Reg.RegexCoreE1Product.IsMatch(Express)) //E1 Product
                {
                    return EnumExpressType.E1Product;
                }
                if (Reg.RegexmVWithBinSearchMode.IsMatch(Express))
                {
                    return EnumExpressType.MvSearch;
                }
                if (Reg.RegexmVWithProductMode.IsMatch(Express))
                {
                    return EnumExpressType.MvProduct;
                }
                if (Reg.RegexAllProduct.IsMatch(Express)) //Product
                {
                    if (Express.Contains("GB", System.StringComparison.OrdinalIgnoreCase))
                    {
                        return EnumExpressType.ProductGb;
                    }

                    return EnumExpressType.Product;
                }
                if (Reg.RegexAllmVWithMode.IsMatch(Express)) //All => +700.25mV (MS001)
                {
                    return EnumExpressType.MvWithMode;
                }
                if (Reg.RegexAllmV.IsMatch(Express)) //All => +700.25mV
                {
                    return EnumExpressType.AllMv;
                }
                if (Reg.RegexBinningVmax.IsMatch(Express) || // Add for CPU_SRAM as core power CPVmax , BinningVmax
                    Reg.RegexCpVmax.IsMatch(Express))
                {
                    return EnumExpressType.VMax;
                }
                return EnumExpressType.Other;
            }
        }

        public BinCutExpress()
        {
        }

        public BinCutExpress(BinCutExpress binCutExpress)
        {
            if (binCutExpress == null)
            {
                return;
            }

            Title = binCutExpress.Title;
            Express = binCutExpress.Express;
        }

        public BinCutExpress Copy()
        {
            return new BinCutExpress(this);
        }

        public string GetMode()
        {
            return BinCutAlgorithmService.GetModeByName(Express);
        }

        public bool IsFunction()
        {
            return Reg.RegexFunction.IsMatch(Express);
        }
    }
}
