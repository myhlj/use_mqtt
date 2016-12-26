#!/usr/bin/env python
#encoding=utf-8
from use_xl import use_xlrd,use_xlwt
import sys,json
from use_mqtt import use_mqtt
import multiprocessing,time
#import urllib2
import requests,os,io,socket
#from PIL import Image

def xlwt_process(queue,lock,use_xlwt,ip,port,dir_name):
    if not os.path.exists('./' + dir_name):
        os.makedirs('./' + dir_name)

    while True:
        try:
            lock.acquire()
            if queue.empty() != True:
                info = queue.get()
                lock.release()
                dict_info = json.loads(info)  
                error = False
                for key in dict_info:
                    if key == 'img_f' or key == 'img_v':
                        url = "http://" + "192.168.30.222" + ":" + str(8081) + dict_info[key]
                        try:
                            data = requests.get(url)#urllib2.urlopen(url).read()
                            jpg_data = data.content

                            file_name = './' + dir_name + '/' + dict_info[key].split('/')[-1] + ".jpg"
                            f = open(file_name,'wb')
                            f.write(jpg_data)
                            f.close()
                            '''
                            file_name = './' + dir_name + '/' + dict_info[key].split('/')[-1] + ".bmp"
                            #jpg 转 bmp
                            image = Image.open(io.BytesIO(jpg_data))
                            bmparr = io.BytesIO()
                            image.save(bmparr,format='BMP')
                            bmp_data = bmparr.getvalue() 
                            #写入文件
                            f = open(file_name,'wb')
                            f.write(bmp_data)
                            f.close()
                            '''
                            dict_info[key] = file_name
                        except (requests.ConnectionError,IOError),e:
                            print e
                            error = True
                            break
                if error:#出现错误,丢弃本条
                    continue

                print '数据:',dict_info

                if len(use_xlwt.dict_sheet_row) == 0:
                    use_xlwt.add_default_sheet()
                    use_xlwt.write_row_content(dict_info)
                    use_xlwt.save_file()
                else:
                    use_xlwt.write_row_content(dict_info)
                    use_xlwt.save_file()
            else:
                lock.release()
                time.sleep(1)
        except KeyboardInterrupt,e:
            print e
            return 0

if __name__ == '__main__':
    try:
        use_xlwt = use_xlwt(sys.argv[1])
        use_xlwt.get_min_rows()#获取输入的excel中可用的sheet和当前sheet的行数,返回dict
    except IndexError,e:
        print '请输入excel文件名,执行格式(./main.py sample.xls 192.168.30.47 1883 Image):',e
        sys.exit()

    #http client
    try:
        ip = sys.argv[2]
        port = int(sys.argv[3])
    except IndexError,e:
        print '请输入ip地址和端口,执行格式(./main.py sample.xls 192.168.30.47 1883 Image):',e
        sys.exit()

    #dirname
    try:
        dir_name = sys.argv[4]
    except IndexError,e:
        print '请输入保存图片文件夹的名称,执行格式(./main.py sample.xls 192.168.30.47 1883 Image):',e

    #多进程
    lock = multiprocessing.Lock()
    queue = multiprocessing.Queue(10)
    #xlwt进程
    process_excel = multiprocessing.Process(target=xlwt_process,args=(queue,lock,use_xlwt,ip,port,dir_name))
    process_excel.start()
    '''
    dict_info = {'time':'2017-12-20','index':2,'score':999,'bmp':'haha'}
    dict_info2 = {'time':'2017-12-20','index':2,'score':999,'jpg':'2333'}
    arr_dict = [dict_info,dict_info2]

    for dict_content in arr_dict:
        print use_xlwt.dict_sheet_row
        if len(use_xlwt.dict_sheet_row) == 0:
            use_xlwt.add_default_sheet()
            use_xlwt.write_row_content(dict_content)
            use_xlwt.save_file()
        else:
            use_xlwt.write_row_content(dict_content)
            use_xlwt.save_file()
    ''' 
    #mqtt进程(主进程)
    use_mqtt = use_mqtt()
    try:
        use_mqtt.connect(ip,port)
    except socket.error,e:
        print 'connect failed!'
        exit()

    use_mqtt.mqttc.user_data_set(queue)
    use_mqtt.settopic('/tx1/00005/#','/tx1/00005')
    use_mqtt.subscribe()
    try:
        use_mqtt.loop_forever()
    except KeyboardInterrupt,e:
        sys.exit()
