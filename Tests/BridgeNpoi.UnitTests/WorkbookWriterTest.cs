using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace AwesomeExcel.BridgeNpoi.UnitTests;

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
