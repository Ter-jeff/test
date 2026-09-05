using System;

using Automation.PreCheck.AllParaData;
using Automation.Static;

namespace Automation.GenerateIgxl
{
    public abstract class WorkFlowBase<T> : IDisposable where T : ParaData
    {
        public abstract bool PreCheckFlow(T paraData);
        public abstract void WorkFlow(T paraData);

        public void Execute(T paraData)
        {
            if (!PreCheckFlow(paraData))
            {
                return;
            }

            WorkFlow(paraData);
        }


        public void Dispose()
        {
            GC.Collect();
            for (int j = 0; j < GC.MaxGeneration; j++)
            {
                GC.Collect(j);
                GC.WaitForPendingFinalizers();
            }
            GC.SuppressFinalize(this);
        }

        public void Print()
        {
            TestProgram.Print();
        }
    }
}
