using AwesomeExcel.BridgeNPOI;

namespace Tests.BridgeNpoi;

[TestClass]
public class WorkbookWriterTest
{
    [TestMethod]
    public void Convert_NullWorkbook_ShouldThrow_ArgumentNullException()
    {
        WorkbookWriter writer = new();
        Assert.ThrowsException<ArgumentNullException>(() => writer.Write(null));
    }
}
