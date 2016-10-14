'
' 日付: 2016/10/09
'
Imports NUnit.Framework

Imports Common.Text

<TestFixture> _
Public Class TestMatchingText
  
  <Test> _
  Public Sub TestWord
    Dim mt1 As New MatchingText("abcd", MatchingMode.Forward)
    Assert.AreEqual("abcd", mt1.Word)
    
    Dim mt2 As New MatchingText("abcd  efgh", MatchingMode.Forward)
    Assert.AreEqual("abcd  efgh", mt2.Word)
    
    Dim mt3 As New MatchingText("", MatchingMode.Forward)
    Assert.AreEqual("", mt3.Word)
  End Sub
  
  <Test> _
  Public Sub TestMatching
    Dim mt1 As New MatchingText("abc", MatchingMode.Forward)
    Dim mt2 As New MatchingText("abcdef", MatchingMode.Perfection)
    Dim mt3 As New MatchingText("abcdef", MatchingMode.Part)
    Dim mt4 As New MatchingText("0abcdef0", MatchingMode.Backward)
    Dim mt5 As New MatchingText("abc0abcdef0", MatchingMode.Forward)
    Assert.True(mt1.Matching(mt2, False))
    Assert.False(mt2.Matching(mt1, False))
    Assert.True(mt2.Matching(mt3, False))
    Assert.True(mt3.Matching(mt4, False))
    Assert.True(mt4.Matching(mt5, False))
    
    Dim mmt1 As New MatchingText("abc  def", MatchingMode.Forward)
    Dim t1 As New MatchingText("abcd", MatchingMode.Perfection)
    Dim t2 As New MatchingText("defg", MatchingMode.Perfection)
    Dim t3 As New MatchingText("c  def", MatchingMode.Perfection)
    Assert.True(mmt1.Matching(t1, False))
    Assert.True(mmt1.Matching(t2, False))
    Assert.False(mmt1.Matching(t3, False))
    
    Dim emp1 As New MatchingText("", MatchingMode.Forward)
    Dim t4 As New MatchingText("a", MatchingMode.Perfection)
    Assert.True(emp1.Matching(t4, False))
    Assert.True(t4.Matching(emp1, True))
    Assert.False(t4.Matching(emp1, False))
    
    
  End Sub
End Class
