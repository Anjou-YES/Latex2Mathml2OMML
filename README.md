# Latex 转 Mathml 转 OMML

     最近在做项目，其中有一个小功能是要将latex格式的公式表达式插入到word中，这个东西整整折磨了我一天，人都麻了...

        为了后续的同学们不再因为这个不算难(只能说比较偏)的问题困扰，此贴用来解决此类转化问题，也算是做个记录，避免遗忘~

        好，正文开始！

首先，我们需要这么几个包

    latex2mathml

    python-docx

    lxml

其中pip安装的话就是                                                                                           

    pip install python-docx lxml latex2mathml
 
## **_Latex --> Mathml_**

    mathml_output = latex2mathml.converter.convert(##传入latex表达式##)`

## Mathml --> OMML
    tree = etree.fromstring(mathml_output)

    xslt = etree.parse('MML2OMML.XSL')

    transform = etree.XSLT(xslt)

    new_dom = transform(tree)

## CSDN链接：

链接：[https://blog.csdn.net/weixin_52654243/article/details/128234365?spm=1001.2014.3001.5501](https://blog.csdn.net/weixin_52654243/article/details/128234365?spm=1001.2014.3001.5501)

## 百度网盘链接：

链接：[https://pan.baidu.com/s/19xvfcQaD3ETiJWPwUbfmPw?pwd=6666]() 

提取码：**6666**
