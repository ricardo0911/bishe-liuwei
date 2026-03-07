from docx import Document
p = "C:\\Users\\root\\Desktop\\刘伟\\刘伟-feedback3.docx"
doc = Document(p)
text = doc.paragraphs[204].text.strip()
print(repr(text))
print(text == '表3-2 商家表')
