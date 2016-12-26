#encoding=utf-8
import paho.mqtt.client as mqtt
import json,socket

class use_mqtt():
    def __init__(self):
        self.mqttc = mqtt.Client()
        self.subscribe_topic = "lj/#"
        self.publish_topic = "lj/"
        self.mqttc.on_connect = self.on_connect
        self.mqttc.on_message = self.on_message
        self.mqttc.on_subscribe = self.on_subscribe

    def on_connect(self,client,userdata,flagc,rc):
        print("Connected with result code" + str(rc))
    
    def on_message(self,client,userdata,msg):
        #print("Received on topic: " + msg.topic + " Message: " + str(msg.payload) + "\n");
        #dict_content = json.loads(msg.payload)
        #print dict_content
        userdata.put(msg.payload)#放入队列

    def on_subscribe(self,mosq,obj,mid,granted_qos):
        print("on_subscribe ok")

    def settopic(self,subscribe_topic,publish_topic):
        self.subscribe_topic = subscribe_topic
        self.publish_topic = publish_topic

    def connect(self,ip,port):
        self.mqttc.connect(ip,port,60)

    def subscribe(self):
        #print self.subscribe_topic
        self.mqttc.subscribe(self.subscribe_topic,0)

    def publish(self,jsoncontent):
        #print self.publish_topic
        self.mqttc.publish(self.publish_topic,jsoncontent)

    def loop_forever(self):
        self.mqttc.loop_forever()


#if __name__ == '__main__':
    #use_mqtt = use_mqtt()
    #try:
        #use_mqtt.connect('127.0.0.1',1883)
    #except socket.error,e:
        #print '链接服务器失败:',e
        #exit()
    #use_mqtt.subscribe()
    #use_mqtt.publish('life is short,enjoy lift')
    #use_mqtt.loop_forever()
