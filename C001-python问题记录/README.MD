### 1.python gui打包,弹出缺少mkl_intel_thread.dll错误
* 解决方法：
* 复制以下dll文件至exe同级目录
 libiomp5md.dll
 mkl_avx.dll
 mkl_avx2.dll
 mkl_avx512.dll
 mkl_core.dll
 mkl_intel_thread.dll
 mkl_p4.dll
 mkl_p4m.dll
 mkl_p4m3.dll
 mkl_rt.dll
 mkl_sequential.dll
 mkl_tbb_thread.dll
 mkl_vml_avx.dll
 mkl_vml_avx2.dll
 mkl_vml_avx512.dll
 mkl_vml_cmpt.dll
 mkl_vml_ia.dll
 mkl_vml_p4.dll
 mkl_vml_p4m.dll
 mkl_vml_p4m2.dll
 mkl_vml_p4m3.dll

### 
matplotlib 画图显示中文字符问题
####1.在新py文件添加以下两行代码，运行，打印出一个文件路径
    *import matplotlib
    *print(matplotlib.matplotlib_fname())
####2.打开文件，在行 #font.sans-serif: DejaVu Sans 添加SimHei
    效果：#font.sans-serif: SimHei, DejaVu Sans
