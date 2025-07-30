## Hello! I'm INSEONG LEE (inseong101)


MACRO CODE
<pre><code>
Sub JPG()
Dim i&, s As ChartObject, 학번$, r As Range
Set r = [A1:AK50]
For i = 1 To 120
  학번 = Sheets("성적통지표").Cells(i, 1)
  If 학번 <> "" Then
    [AG2:AH3] = 학번
    r.CopyPicture xlPrinter, xlPicture
    Application.Wait Now + TimeValue("0:00:01")
    Set s = ActiveSheet.ChartObjects.Add(0, 0, r.Width * 3, r.Height * 3)
    s.Chart.Paste: s.Chart.Export "C:\Users\김문주\OneDrive - pusan.ac.kr\바탕 화면\성적표캡차\" & 학번 & ".png", "PNG": s.Delete
  End If
Next
End Sub
</code></pre>
MACRO CODE

[![Google Scholar Badge](https://img.shields.io/badge/-Google%20Scholar-4285F4?style=flat-square&logo=Google-Scholar&logoColor=white&link=https://scholar.google.com/citations?user=GeOAGbwAAAAJ)](https://scholar.google.com/citations?user=GeOAGbwAAAAJ) [![ORCID Badge](https://img.shields.io/badge/-ORCID-A6CE39?style=flat-square&logo=ORCID&logoColor=white&link=https://orcid.org/0000-0002-7423-0090)](https://orcid.org/0000-0001-7445-3983) [![GitHub.io Badge](https://img.shields.io/badge/-GitHub.io-181717?style=flat-square&logo=GitHub&logoColor=white&link=https://inseong101.github.io)](https://inseong101.github.io) [![Naver Blog Badge](https://img.shields.io/badge/-Naver%20Blog-03C75A?style=flat-square&logo=Naver&logoColor=white&link=https://blog.naver.com/pnu_kmed)](https://blog.naver.com/pnu_kmed)


- ⚛️ B.Sc. Physics, SKKU
- 🌿 M.S. Korean Medicine, PNU 
- 💻 Beginner in Coidng
- 📈 I enjoy learning new technology.
