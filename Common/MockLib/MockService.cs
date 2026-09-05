using System;

using CommonLib.Static;

using Moq;

namespace MockLib
{
    public static class MockService
    {
        public static readonly string KeyName = @"SOFTWARE\MyUnitTestApp";
        public static readonly string ValueName = "IsUnitTest";
        public static readonly DateTime DefaultDate = new(2022, 01, 01, 00, 00, 00, DateTimeKind.Utc);

        public static void Mock()
        {
            TimeMock();

            AssemblyMock();
        }

        private static void TimeMock()
        {
            TimeContext.ResetToDefault();
            var timeMock = new Mock<TimeProvider>();
            timeMock.Setup(tp => tp.GetUtcNow()).Returns(DefaultDate);
            timeMock.SetupGet(tp => tp.LocalTimeZone).Returns(TimeZoneInfo.Utc);
            TimeContext.Current = timeMock.Object;
        }

        private static void AssemblyMock()
        {
            AssemblyProvider.Current = MockAssemblyProvider.Instance;
        }
    }
}
