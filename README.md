# Latex 转 Mathml 转 OMML

     最近在做项目，其中有一个小功能是要将latex格式的公式表达式插入到word中，这个东西整整折磨了我一天，人都麻了...

        为了后续的同学们不再因为这个不算难(只能说比较偏)的问题困扰，此贴用来解决此类转化问题，也算是做个记录，避免遗忘~

        好，正文开始！

首先，我们需要这么几个包

`import latex2mathml.converter`

`from docx import shared`

`from docx.oxml.ns import qn`

`from docx import Document`

`from lxml import etree`

其中pip安装的话就是                                                                                           

`pip install python-docx lxml latex2mathml`
 
## **_Latex --> Mathml_**

`#latex表达式

latex_input = '\exp\left[\int d^{4}x g\phi\bar{\psi}\psi\right]=\sum_{n=0}^{\infty}\frac{g^{n}}{n!}\left(\int d^{4}x\phi\bar{\psi}\psi\right)^{n}.'
 
#mathml格式

mathml_output = latex2mathml.converter.convert(latex_input)`

而mathml转office word（docx格式）可以识别的不是就是之前让我带上痛苦面具的部分了，这边需要MS office的MML2OMML.XSL文件（一般在MS Office的文件夹的root/office xx文件夹下，下面会放网盘链接，需要自取）

## Mathml --> OMML
`#MML2OMML.XSL
tree = etree.fromstring(mathml_output)

xslt = etree.parse('MML2OMML.XSL')

transform = etree.XSLT(xslt)

#new_dom就是omml格式啦

new_dom = transform(tree)`


下一步就是将omml写入到word了，这一步也有一个小坑，要用_element的方法去将公式append到word的p字段中，代码如下：


`doc = Document()

#定义英文及数字文字字体

doc.styles['Normal'].font.name = 'Times New Roman'

#定义中文文字字体

doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

#定义字体大小

doc.styles['Normal'].font.size = shared.Pt(9)

paragraph = doc.add_paragraph(style=None)`

#没错，就是这步

paragraph._element.append(new_dom.getroot())

### 上面就已经将公式插入到了word中了，下面附上完整代码

`import latex2mathml.converter

from docx import shared

from docx.oxml.ns import qn

from docx import Document

from lxml import etree

 
latex_input = "{\\dot{\\epsilon}}\\perp{\\dot{a}},{\\dot{\\epsilon}}\\perp{\\dot{b}}\\,,"

'\exp\left[\int d^{4}x g\phi\bar{\psi}\psi\right]=\sum_{n=0}^{\infty}\frac{g^{n}}{n!}\left(\int d^{4}x\phi\bar{\psi}\psi\right)^{n}.'

mathml_output = latex2mathml.converter.convert(latex_input)

#MML2OMML.XSL

tree = etree.fromstring(mathml_output)

xslt = etree.parse('MML2OMML.XSL')

transform = etree.XSLT(xslt)

new_dom = transform(tree)
 
doc = Document()

#定义英文及数字文字字体

doc.styles['Normal'].font.name = 'Times New Roman'

#定义中文文字字体

doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

#定义字体大小

doc.styles['Normal'].font.size = shared.Pt(9)

paragraph = doc.add_paragraph(style=None)

paragraph._element.append(new_dom.getroot())

docx_path = '{}.docx'.format('test')

doc.save(docx_path)`


ok~大功告成！

#### 百度网盘链接：

链接：[https://pan.baidu.com/s/19xvfcQaD3ETiJWPwUbfmPw?pwd=6666]() 

提取码：**6666**
