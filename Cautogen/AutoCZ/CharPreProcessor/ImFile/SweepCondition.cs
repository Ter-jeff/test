using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.ImFile
{
    public class SweepCondition
    {
        #region properties
        public string XSupplyP = "";

        public string TrackSupplyP = "";

        public string YSupplyP = "";

        public string YTrackSupplyP = "";

        public string ZSupplyP = "";

        public string ZTrackSupplyP = "";

        public List<SweepPoint> XSweepList = new List<SweepPoint>();

        public List<SweepPoint> YSweepList = new List<SweepPoint>();

        public List<SweepPoint> ZSweepList = new List<SweepPoint>();

        private readonly Characterization _charRow;

        private readonly List<string> _primaryNameX = new List<string>();

        private readonly List<string> _primaryNameY = new List<string>();

        private readonly List<string> _primaryNameZ = new List<string>();

        private readonly List<string> _trackingShmooNamesX = new List<string>();

        private readonly List<string> _trackingShmooNamesY = new List<string>();

        private readonly List<string> _trackingShmooNamesZ = new List<string>();

        private string _xsupply = "";

        private string _xtsupply = "";

        private string _ysupply = "";

        private string _ytsupply = "";

        private string _zsupply = "";

        private string _ztsupply = "";

        public string SetupName
        {
            get
            {
                if (_xsupply + _ysupply + _xtsupply + XSupplyP == "")
                {
                    return "";
                }

                string xyStr = "";
                if (XSweepList.Count > 0 && YSweepList.Count > 0)
                {
                    xyStr = "_VS";
                }

                string zStr = "";
                if ((XSweepList.Count > 0 || YSweepList.Count > 0) && ZSweepList.Count > 0)
                {
                    zStr = "_VS";
                }

                if (_xtsupply != "")
                {
                    _xtsupply = "_T" + _xtsupply;
                }

                if (XSupplyP != "" && _xsupply != "")
                {
                    XSupplyP = "_T" + XSupplyP;
                }

                if (TrackSupplyP != "")
                {
                    TrackSupplyP = "_T" + TrackSupplyP;
                }

                if (YSupplyP != "" && _ysupply != "")
                {
                    YSupplyP = "_T" + YSupplyP;
                }

                if (YTrackSupplyP != "")
                {
                    YTrackSupplyP = "_T" + YTrackSupplyP;
                }

                if (ZSupplyP != "" && _ysupply != "")
                {
                    ZSupplyP = "_T" + ZSupplyP;
                }

                if (ZTrackSupplyP != "")
                {
                    ZTrackSupplyP = "_T" + ZTrackSupplyP;
                }

                return _charRow.Category +
                    _xsupply + _xtsupply + XSupplyP + TrackSupplyP + xyStr +
                    _ysupply + _ytsupply + YSupplyP + YTrackSupplyP + zStr +
                    _zsupply + _ztsupply + ZSupplyP + ZTrackSupplyP +
                    "_" + _charRow.Search;
            }
        }

        #endregion

        #region constructor

        public SweepCondition(Characterization charRow)
        {
            _charRow = charRow;

            foreach (ShmooSpec spec in _charRow.PrimaryShmooXList)
            {
                _primaryNameX.Add(spec.Name);
            }

            foreach (ShmooSpec spec in _charRow.PrimaryShmooYList)
            {
                _primaryNameY.Add(spec.Name);
            }

            foreach (ShmooSpec spec in _charRow.PrimaryShmooZList)
            {
                _primaryNameZ.Add(spec.Name);
            }

            foreach (ShmooSpec spec in _charRow.TrackingSpecX)
            {
                _trackingShmooNamesX.Add(spec.Name);
            }

            foreach (ShmooSpec spec in _charRow.TrackingSpecY)
            {
                _trackingShmooNamesY.Add(spec.Name);
            }

            foreach (ShmooSpec spec in _charRow.TrackingSpecZ)
            {
                _trackingShmooNamesZ.Add(spec.Name);
            }

            _WritePowerSupplyX(charRow);
            _WritePowerSupplyY(charRow);
            _WritePowerSupplyZ(charRow);
            _WritePinSweep(charRow);
        }
        #endregion

        #region methods

        private void _WritePowerSupplyX(Characterization charRow)
        {
            string primaryNamesX = "";
            string levelLabel = charRow.GetLevelLabel("x");

            // primary
            foreach (ShmooSpec spec in charRow.PrimaryShmooXList)
            {
                _xsupply += "_" + spec.Name;
                primaryNamesX += "," + spec.Name;
            }

            if (charRow.PrimaryShmooXList.Count > 0)
            {
                ShmooSpec spec = charRow.PrimaryShmooXList[0];
                XSweepList.Add(new SweepPoint
                {
                    SweepName = levelLabel + ":" + primaryNamesX.Trim(','),
                    Start = spec.Start,
                    Stop = spec.Stop,
                    Step = spec.Step
                });
            }

            // tracking
            foreach (ShmooSpec spec in charRow.TrackingSpecX)
            {
                if (spec.IsPowerPin) //VDD shmoo 
                {
                    _xtsupply += "_" + spec.Name;

                    string trackingNames1X = spec.Name;
                    XSweepList.Add(new SweepPoint
                    {
                        SweepName = levelLabel + ":" + trackingNames1X,
                        Start = spec.Start,
                        Stop = spec.Stop,
                        Step = spec.Step
                    });
                }

                else if (Regex.IsMatch(spec.Name, "^REF.*")) //REF Freq
                {
                    if (_trackingShmooNamesX.Contains(spec.Name))
                    {
                        _xtsupply += "_XIO_Freq";
                    }
                    else
                    {
                        _xsupply += "_XIO_Freq";
                        XSweepList.Add(new SweepPoint
                        {
                            SweepName = "AC Spec:XIO_Freq",
                            Start = spec.Start,
                            Stop = spec.Stop,
                            Step = spec.Step
                        });
                    }
                }
                else //Shift Freq
                {
                    if (_trackingShmooNamesX.Contains(spec.Name))
                    {
                        _xtsupply += "_TCK_Freq";
                    }
                    else
                    {
                        _xsupply += "_TCK_Freq";
                        XSweepList.Add(new SweepPoint
                        {
                            SweepName = "AC Spec:TCK_Freq",
                            Start = spec.Start,
                            Stop = spec.Stop,
                            Step = spec.Step
                        });
                    }
                }
            }
        }

        private void _WritePowerSupplyY(Characterization charRow)
        {
            bool firstPrimaryY = true;
            string primaryNamesY = "";
            string levelLabel = charRow.GetLevelLabel("y");

            foreach (ShmooSpec shmoo in charRow.PowerSupplyY)
            {

                if (shmoo.IsPowerPin) //VDD shmoo 
                {
                    if (shmoo.Start == shmoo.Stop || shmoo.Stop == "")
                    {
                        continue;
                    }

                    if (_trackingShmooNamesY.Contains(shmoo.Name))
                    {
                        _ytsupply += "_" + shmoo.Name;

                        YSweepList.Add(new SweepPoint
                        {
                            SweepName = levelLabel + ":" + shmoo.Name,
                            Start = shmoo.Start,
                            Stop = shmoo.Stop,
                            Step = shmoo.Step,
                        });

                    }
                    else if (_primaryNameY.Contains(shmoo.Name))
                    {
                        _ysupply += "_" + shmoo.Name;

                        primaryNamesY += "," + shmoo.Name;

                        if (firstPrimaryY)
                        {
                            YSweepList.Add(new SweepPoint
                            {
                                SweepName = levelLabel + ":" + primaryNamesY.Trim(','),
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                        firstPrimaryY = false;
                    }
                    else
                    {
                        const string outString = "NumStepNotMatch";
                        ErrorManager.AddError(ErrorType.ErrorShmooRange, charRow.SheetName, charRow.RowNum,
                            shmoo.ColIndex, charRow.Use,
                            outString);
                        ErrorReportManager.AddError(CharErrorType.E_ErrorShmooRange_02, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, []);
                    }
                }
                else
                {
                    if (shmoo.Start == shmoo.Stop || shmoo.Stop == "")
                    {
                        continue;
                    }

                    if (Regex.IsMatch(shmoo.Name, "XI0", RegexOptions.IgnoreCase)) //REF Freq
                    {
                        if (_trackingShmooNamesY.Contains(shmoo.Name))
                        {
                            _ytsupply += "_XIO_Freq";
                        }
                        else
                        {
                            _ysupply += "_XIO_Freq";
                            YSweepList.Add(new SweepPoint
                            {
                                SweepName = "Ac Spec:XIO_Freq",
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                    }
                    else //Shift Freq
                    {
                        if (_trackingShmooNamesY.Contains(shmoo.Name))
                        {
                            _ytsupply += "_TCK_Freq";
                        }
                        else
                        {
                            _ysupply += "_TCK_Freq";
                            YSweepList.Add(new SweepPoint
                            {
                                SweepName = "Ac Spec:TCK_Freq",
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                    }
                }
            }
        }

        private void _WritePowerSupplyZ(Characterization charRow)
        {
            bool firstPrimaryZ = true;
            string primaryNamesZ = "";
            string levelLabel = charRow.GetLevelLabel("z");

            foreach (ShmooSpec shmoo in charRow.PowerSupplyZ)
            {
                if (shmoo.IsPowerPin) //VDD shmoo 
                {
                    if (shmoo.Start == shmoo.Stop || shmoo.Stop == "")
                    {
                        continue;
                    }

                    if (_trackingShmooNamesZ.Contains(shmoo.Name))
                    {
                        _ztsupply += "_" + shmoo.Name;

                        ZSweepList.Add(new SweepPoint
                        {
                            SweepName = levelLabel + ":" + shmoo.Name,
                            Start = shmoo.Start,
                            Stop = shmoo.Stop,
                            Step = shmoo.Step,
                        });
                    }
                    else if (_primaryNameZ.Contains(shmoo.Name))
                    {
                        _zsupply += "_" + shmoo.Name;

                        primaryNamesZ += "," + shmoo.Name;

                        if (firstPrimaryZ)
                        {
                            ZSweepList.Add(new SweepPoint
                            {
                                SweepName = levelLabel + ":" + primaryNamesZ.Trim(','),
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                        firstPrimaryZ = false;
                    }
                    else
                    {
                        const string outString = "NumStepNotMatch";
                        ErrorManager.AddError(ErrorType.ErrorShmooRange, charRow.SheetName, charRow.RowNum,
                            shmoo.ColIndex, charRow.Use,
                            outString);
                        ErrorReportManager.AddError(CharErrorType.E_ErrorShmooRange_02, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, []);
                    }
                }
                else
                {
                    if (shmoo.Start == shmoo.Stop || shmoo.Stop == "")
                    {
                        continue;
                    }

                    if (Regex.IsMatch(shmoo.Name, "^REF.*")) //REF Freq
                    {
                        if (_trackingShmooNamesZ.Contains(shmoo.Name))
                        {
                            _ztsupply += "_XIO_Freq";
                        }
                        else
                        {
                            _zsupply += "_XIO_Freq";
                            ZSweepList.Add(new SweepPoint
                            {
                                SweepName = "Ac Spec:XIO_Freq",
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                    }
                    else //Shift Freq
                    {
                        if (_trackingShmooNamesZ.Contains(shmoo.Name))
                        {
                            _ztsupply += "_TCK_Freq";
                        }
                        else
                        {
                            _zsupply += "_TCK_Freq";
                            ZSweepList.Add(new SweepPoint
                            {
                                SweepName = "Ac Spec:TCK_Freq",
                                Start = shmoo.Start,
                                Stop = shmoo.Stop,
                                Step = shmoo.Step,
                            });
                        }
                    }
                }
            }
        }

        private void _WritePinSweep(Characterization charRow)
        {
            var xsupplyPin = new Pin();

            foreach (Pin pin in charRow.Pins.Where(pin => (pin.Start != "" && pin.Stop != "" && pin.Start != pin.Stop)
                || pin.PinType.Equals("AC Spec", StringComparison.OrdinalIgnoreCase)))
            {
                if (pin.Order.IndexOf("Y", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    //
                    if (YSupplyP == "")
                    {
                        xsupplyPin = pin;
                        YSupplyP = "_" + pin.PinName + "_" + pin.PinType;
                        YSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                        if (xsupplyPin.Step == "")
                        {
                            xsupplyPin.Step = "0.005";
                        }
                    }
                    else
                    {
                        double primaryStepNum = (Convert.ToDouble(xsupplyPin.Start) - Convert.ToDouble(xsupplyPin.Stop)) /
                                             Convert.ToDouble(xsupplyPin.Step);

                        double step = Convert.ToDouble(pin.Step == "" ? xsupplyPin.Step : pin.Step);

                        double otherStepNum = (Convert.ToDouble(pin.Start) - Convert.ToDouble(pin.Stop)) / step;

                        if (!(Math.Abs(primaryStepNum - otherStepNum) < 1E-12))
                        {
                            continue;
                        }

                        YTrackSupplyP += "_" + pin.PinName + "_" + pin.PinType;

                        YSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                    }
                    //
                }
                else if (pin.Order.IndexOf("Z", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    //
                    if (ZSupplyP == "")
                    {
                        xsupplyPin = pin;
                        ZSupplyP = "_" + pin.PinName + "_" + pin.PinType;
                        ZSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                        if (xsupplyPin.Step == "")
                        {
                            xsupplyPin.Step = "0.005";
                        }
                    }
                    else
                    {
                        double primaryStepNum = (Convert.ToDouble(xsupplyPin.Start) - Convert.ToDouble(xsupplyPin.Stop)) /
                                             Convert.ToDouble(xsupplyPin.Step);

                        double step = Convert.ToDouble(pin.Step == "" ? xsupplyPin.Step : pin.Step);

                        double otherStepNum = (Convert.ToDouble(pin.Start) - Convert.ToDouble(pin.Stop)) / step;

                        if (!(Math.Abs(primaryStepNum - otherStepNum) < 1E-12))
                        {
                            continue;
                        }

                        ZTrackSupplyP += "_" + pin.PinName + "_" + pin.PinType;

                        ZSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                    }
                    //
                }
                else
                {
                    if (XSupplyP == "")
                    {
                        xsupplyPin = pin;
                        XSupplyP = "_" + pin.PinName + "_" + pin.PinType;
                        XSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                        if (xsupplyPin.Step == "")
                        {
                            xsupplyPin.Step = "0.005";
                        }
                    }
                    else
                    {
                        double primaryStepNum = (Convert.ToDouble(xsupplyPin.Start) - Convert.ToDouble(xsupplyPin.Stop)) /
                                             Convert.ToDouble(xsupplyPin.Step);

                        double step = Convert.ToDouble(pin.Step == "" ? xsupplyPin.Step : pin.Step);

                        double otherStepNum = (Convert.ToDouble(pin.Start) - Convert.ToDouble(pin.Stop)) / step;

                        if (!(Math.Abs(primaryStepNum - otherStepNum) < 1E-12))
                        {
                            continue;
                        }

                        TrackSupplyP += "_" + pin.PinName + "_" + pin.PinType;

                        XSweepList.Add(new SweepPoint
                        {
                            SweepName = pin.PinType + ":" + pin.PinName,
                            Start = pin.Start,
                            Stop = pin.Stop,
                            Step = pin.Step,
                        });
                    }
                }
            }
        }

        #endregion
    }
}
