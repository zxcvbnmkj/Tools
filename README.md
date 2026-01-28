## dir_tree
自动生成当前工程的文件树，已在 windows 和 mac 中测试过，都可正常使用。还未在 linux 测试。

使用命令行指定特定路径的方法是直接在文件名后面加路径
```
python dir_tree.py /Users/zl/sphinx_demo
```
## md2word
- 点击[链接](https://github.com/jgm/pandoc/blob/main/INSTALL.md)下载 pandoc 软件，并将该软件的路径 `C:\Program Files\Pandoc` 添加到环境变量中。最后请安装以下两个 python 库：
```
pip install pypandoc
pip install python-docx
```
- 如果遇到以下报错，是因为 word 文件当前处于打开状态，把它关闭后再次运行即可
```txt
Traceback (most recent call last):
  File "E:\feng\md2word.py", line 75, in <module>
    convert_md_to_docx(args.input_file, args.use_temp)
  File "E:\feng\md2word.py", line 64, in convert_md_to_docx
    pypandoc.convert_file(input_file, 'docx', outputfile=output_file, extra_args=extra_args)
  File "E:\1sw\miniconda3\envs\py310cu124\lib\site-packages\pypandoc\__init__.py", line 255, in convert_file
    return _convert_input(
  File "E:\1sw\miniconda3\envs\py310cu124\lib\site-packages\pypandoc\__init__.py", line 528, in _convert_input
    raise RuntimeError(
```
- todo:
  - 标题的字体有点问题，文字会一大一小的。但是代码设置的标题字体根本没有起到左右，但设置的颜色（黑色）生效了，原先部分标题是蓝色
  <img width="589" height="211" alt="2f122fea2b39482cca29c9ec8af1144" src="https://github.com/user-attachments/assets/07774ade-4e78-4087-8f9d-504a0d33fcd1" />
  
  - 代码块没有特殊处理，在 word 中很难看
