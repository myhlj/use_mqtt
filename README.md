用于测试mqtt_client订阅信息和xlwt写入信息到xml文件的python脚本
========================================
# 该测试脚本适用于python2.7
****
# 使用方法
    linux下先确保安装python2.7,可以用python --version禅看版本信息
    安装pip install paho-mqtt,pip install xlrd,pip install xlwt,pip install xlutils,pip install requests
    然后执行脚本格式如下:
        python main.py sample.xls 192.168.30.47 1883 Image
        sample.xls为存贮excel文件名 ip地址,端口,Image为保存图片的文件夹名称(也可跟路径)
    结束:ctrl+c
****
# 注意写入到excel中的图片为链接。查看图片需要点击链接
