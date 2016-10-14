'
' 日付: 2016/10/10
'
Imports NUnit.Framework
Imports Common.Util

<TestFixture> _
Public Class TestTextUtils
  <Test> _
  Public Sub TestToCharCode
    Assert.AreEqual(202, TextUtils.ToCharCode("伏見", 3))
    Assert.AreEqual(402, TextUtils.ToCharCode("鴨川", 3))
  End Sub
End Class
