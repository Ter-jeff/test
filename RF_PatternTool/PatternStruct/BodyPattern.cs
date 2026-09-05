namespace RF_PatternTool.PatternStruct
{
    public class BodyPattern
    {
        public List<string> RawDataList = new List<string>();
        public BodyPattern() { }

        public void ReadTemplate()
        {
            List<string> lines = Template.TemplateSet.BodyTemp.Replace("\r", "").Split('\n').ToList();

            foreach (string line in lines)
            {
                if (line.Contains(">"))
                {
                    RawDataList.Add(line.Split('/')[0]);
                }
                else
                {
                    RawDataList.Add(line);
                }
            }
        }

        public string WriteMerge(TextWriter sw, string pattern, VectorInfo info, bool isSource, bool isCap, bool isMeasCall, HashSet<string> matchLoopSAddressSRM)
        {
            string lastVec = "";
            string cappin = "JTAG_TDO";
            string srcpin = "JTAG_TDI";
            #region Writepattern pre-header
            sw.WriteLine("digital_inst = hsdmq;");
            sw.WriteLine("opcode_mode = single;");
            sw.WriteLine("import tset tsetJTAG;");

            if (isMeasCall)
            {
                sw.WriteLine("import subr Global_TEOPEARTION_meas;");
            }
            if (matchLoopSAddressSRM.Count > 0)
            {
                foreach (string loop in matchLoopSAddressSRM)
                {
                    sw.WriteLine("import subr {0};", loop);
                }
            }
            sw.WriteLine("instruments = {");
            sw.WriteLine("({0}):DigCap 1:lsb:serial:auto_trig_enable;", cappin);
            sw.WriteLine("({0}):DigSrc 1:lsb:serial;", srcpin);
            sw.WriteLine("}");
            #endregion
            foreach (string data in RawDataList)
            {
                if (data.Contains("body"))
                {
                    sw.WriteLine(data.Replace("body", pattern));
                }
                else
                {
                    if (data.Contains(">"))
                    {
                        lastVec = data;
                        if (isSource)
                        {
                            sw.WriteLine("(({0}):DigSrc = Start)", srcpin);
                            sw.WriteLine(lastVec.Replace("1", "X").Replace("0", "X"));
                            sw.WriteLine("{0} {1}", "repeat 144", lastVec.Replace("1", "X").Replace("0", "X"));
                            isSource = false;
                            continue;
                        }

                        //// 0 V:0 C:0
                        string patrow = lastVec;
                        if (PatternGenBusiness.IsAddComment)
                        {
                            patrow += info.GetVectorInfo();
                            info.Update();
                        }
                        sw.WriteLine(patrow);
                    }
                    else
                    {
                        sw.WriteLine(data);
                    }
                }
            }
            return lastVec;
        }
    }
}
