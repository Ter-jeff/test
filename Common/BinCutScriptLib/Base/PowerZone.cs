using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using BinCutScriptLib.Printer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Base
{
    [DebuggerDisplay("{Mode}")]
    public class PowerZone
    {
        //VDD_SOC_MS001
        public string PinMode = string.Empty;
        //VDD_SOC
        public string Pin = string.Empty;
        //MS003
        public string Mode = string.Empty;
        //0:linear(step up), 1:ids(step down)
        public EnumIdsMode IdsMode = EnumIdsMode.Linear;
        //Maybe Efuse then real Ids
        public double IdsValue;
        public double IdsValueReal;
        public double IdsValuePowerBinning;
        public int Step = -1;
        public int LastFinalStep = -1;
        private int _finalStep = -1;
        public int FinalStep
        {
            get
            {
                return _finalStep;
            }
            set
            {
                _finalStep = value;
                if (_finalStep != -1 && _finalStep < PossibleSteps.Count)
                {
                    Bin = PossibleSteps[_finalStep].Bin;
                }
            }
        }

        public int StartStep = -1;
        public int StopStep = -1;
        public List<PowerStep> AllSteps = [];
        public EnumPatPf IsLastPatPass
        {
            get
            {
                if (AllSteps.Any(x => x.IsPatternFail))
                {
                    return EnumPatPf.Fail;
                }

                return EnumPatPf.Pass;
            }
        }                             //init = -1, fail for 0, pass for 1
        public bool IsInterplation;
        public int Bin = 1;
        //Turn on flag when gradeSearch fail , Get voltage always is 0 (before limit)
        public bool IsFail;
        //Turn on flag when gradeSearch fail , Skip gradeSearch.(after limit)
        public bool IsBinOut;
        //Turn on flag after adjustVddbinning.
        public bool IsAdjust;
        public EnumSearchStatus SearchStatus { set; get; } = EnumSearchStatus.Init;
        public bool MonoAdjust;

        public PowerStep CurrentStep
        {
            get
            {
                if (Step != -1)
                {
                    return PossibleSteps[Step];
                }

                return null!;
            }
        }
        public List<PowerStep> PossibleSteps
        {
            get
            {
                //for new Interpolation Monotonicity only(interpolation : O && adjust : O)
                if (IsAdjust)
                {
                    int firstStep = AllSteps.FindIndex(x => !x.IsInterpolationDelete && !x.IdsOut);
                    var tmpStep = new List<PowerStep>();
                    for (int i = firstStep; i < AllSteps.Count; i++)
                    {
                        tmpStep.Add(AllSteps[i]);
                    }
                    return tmpStep;
                }
                //original 
                List<PowerStep> steps = AllSteps.FindAll(x => !x.IsInterpolationDelete && !x.IdsOut);
                if (steps.Count == 0 && BinCutConfig.IsDoAll)
                {
                    if (AllSteps.Count != 0)
                    {
                        steps.Add(AllSteps.Last());
                    }
                }
                return steps;
            }
        }
        public PowerZone? AllowPwrRef { get; set; } // link to reference pmode

        public PowerZone()
        {
            if (BinCutConfig.IsDebugPrint)
            {
                _finalStep = 0;
            }
        }

        public PowerZone(PowerZone powerZone)
        {
            if (powerZone == null)
            {
                return;
            }

            PinMode = powerZone.PinMode;
            Pin = powerZone.Pin;
            Mode = powerZone.Mode;
            IdsMode = powerZone.IdsMode;
            IdsValue = powerZone.IdsValue;
            IdsValueReal = powerZone.IdsValueReal;
            IdsValuePowerBinning = powerZone.IdsValuePowerBinning;
            Step = powerZone.Step;
            LastFinalStep = powerZone.LastFinalStep;
            // Copy the backing field directly (not via the FinalStep property setter,
            // which has a side effect of recomputing Bin from PossibleSteps).
            _finalStep = powerZone._finalStep;
            StartStep = powerZone.StartStep;
            StopStep = powerZone.StopStep;
            AllSteps = [.. powerZone.AllSteps.Select(x => x.Copy())];
            IsInterplation = powerZone.IsInterplation;
            Bin = powerZone.Bin;
            IsFail = powerZone.IsFail;
            IsBinOut = powerZone.IsBinOut;
            IsAdjust = powerZone.IsAdjust;
            SearchStatus = powerZone.SearchStatus;
            MonoAdjust = powerZone.MonoAdjust;
            AllowPwrRef = powerZone.AllowPwrRef?.Copy();
        }

        public PowerZone Copy()
        {
            return new PowerZone(this);
        }

        public void SetBin4CandStep()
        {
            int binStep = PossibleSteps.Exists(x => x.Bin == 1 && x.EqName == 1 && string.IsNullOrEmpty(x.EqnBin)) ?
                    PossibleSteps.FindIndex(x => x.Bin == 1 && x.EqName == 1) : PossibleSteps.FindLastIndex(x => x.EqnBin.EqualsIgnoreCase("BIN1"));
            if (binStep != -1)
            {
                Bin = 1;
                Step = binStep;
                StartStep = binStep;
                FinalStep = binStep;
                StopStep = AllSteps.Count - 1;
                SearchStatus = EnumSearchStatus.Search;
            }
        }

        public void SetAdjustTest(bool adjustMode)
        {

            if (adjustMode)
            {
                double lvcc = PossibleSteps[FinalStep].Lvcc;
                int eqn = PossibleSteps[FinalStep].EqName;
                int adjStep = AllSteps.FindIndex(x => x.Lvcc.Equals(lvcc) && x.EqName.Equals(eqn));
                for (int i = 0; i < AllSteps.Count; i++)
                {
                    AllSteps[i].IsInterpolationDelete = i < adjStep;
                }
                FinalStep = 0;
            }
            IsAdjust = adjustMode;
        }

        public int GetPosCount()
        {
            return AllSteps.FindAll(x => !x.IsInterpolationDelete && !x.IdsOut).Count;
        }

        public double GetFinalProductValue(StreamWriter streamWriter)
        {
            try
            {
                if (AllowPwrRef != null)
                {
                    return AllowPwrRef.GetFinalLvcc() + AllSteps.Find(x => x.Bin.Equals(Bin))!.Cp1Gb;
                }

                return PowerZoneHelpers.CalculateFinalValue(this);
            }
            catch
            {
                BinCutPrint.PrintFinalStepNegativeOneErrorMessage(streamWriter, Mode);
                return 9999999;
            }
        }

        public double GetFinalLvcc()
        {
            if (AllowPwrRef != null)
            {
                return AllowPwrRef.GetFinalLvcc();
            }

            if (FinalStep == -1)
            {
                throw new Exception($"FinalStep of {Mode} can not be -1 !!!");
            }

            return PossibleSteps[FinalStep].Lvcc;
        }

        public int GetFinalBin()
        {
            if (SearchStatus != EnumSearchStatus.Search)
            {
                return 0;
            }

            return PossibleSteps[FinalStep].Bin;
        }

        public double GetMonotonicityOffset()
        {
            if (FinalStep == -1)
            {
                return PossibleSteps[0].MonotonicityOffset;
            }

            return PossibleSteps[FinalStep].MonotonicityOffset;
        }

        public int GetFinalEqName()
        {
            if (SearchStatus != EnumSearchStatus.Search)
            {
                return 0;
            }

            return PossibleSteps[FinalStep].EqName;
        }
    }
}
