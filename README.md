## Hello! I'm INSEONG LEE (inseong101)

<pre><code>
Sub JPG()
Dim i&, s As ChartObject, 학번$, sht As Worksheet, tgt As Worksheet, r As Range
Set sht = Sheets("성적통지표"): Set tgt = ActiveSheet: Set r = tgt.Range("A1:AK50")
For i = 1 To 120
  학번 = sht.Cells(i, 1)  ' 1열이란 뜻입니다. A1부터 A120까지 학번 붙여넣고 하얀글자로 숨기세요
  If 학번 <> "" Then
    tgt.Range("AG2:AH3") = 학번 '학번이 들어갈 셀을 안에 넣으세요. 통합셀이면 범위로 넣으세요
    r.CopyPicture xlPrinter, xlPicture
    Application.Wait Now + TimeValue("0:00:01")
    Set s = tgt.ChartObjects.Add(0, 0, r.Width * 3, r.Height * 3)
    s.Chart.ChartArea.Select: s.Chart.Paste
    s.Chart.Export "C:\2. 모의고사 결과분석\성적표캡쳐\" & 학번 & ".png", "PNG" '120장의 파일을 자동 저장할 곳을 경로 복사 하세요 맨 끝에 \붙이셔야 그 폴더 안으로 저장됩니다
    s.Delete
  End If
Next
End Sub
</code></pre>

[![Google Scholar Badge](https://img.shields.io/badge/-Google%20Scholar-4285F4?style=flat-square&logo=Google-Scholar&logoColor=white&link=https://scholar.google.com/citations?user=GeOAGbwAAAAJ)](https://scholar.google.com/citations?user=GeOAGbwAAAAJ) [![ORCID Badge](https://img.shields.io/badge/-ORCID-A6CE39?style=flat-square&logo=ORCID&logoColor=white&link=https://orcid.org/0000-0002-7423-0090)](https://orcid.org/0000-0001-7445-3983) [![GitHub.io Badge](https://img.shields.io/badge/-GitHub.io-181717?style=flat-square&logo=GitHub&logoColor=white&link=https://inseong101.github.io)](https://inseong101.github.io) [![Naver Blog Badge](https://img.shields.io/badge/-Naver%20Blog-03C75A?style=flat-square&logo=Naver&logoColor=white&link=https://blog.naver.com/pnu_kmed)](https://blog.naver.com/pnu_kmed)


- ⚛️ B.Sc. Physics, SKKU
- 🌿 M.S. Korean Medicine, PNU 
- 💻 Beginner in Coidng
- 📈 I enjoy learning new technology.
