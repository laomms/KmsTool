# KMS Local Activation Tool  

#### support win7/8.1/10/server2022 office2010~office2021

![image](https://github.com/laomms/KmsTool/blob/main/kms.JPG)     
![image](https://github.com/laomms/KmsTool/blob/main/kms2.png)   

KMS激活逃不过KMS连接代理:SppExtComObj.Exe这个软件,这个工具在系统目录,验证KMS和合法性全部通过这个,只要对这个可执行文件中的RpcStringBindingComposeW和RpcBindingFromStringBindingW进行拦截,让他将KMS服务器指向本地就可以,这样可以在注册表中搭建一个KMS服务器,实现本地激活(早期Windows版本没有这个工具,直接在注册版模拟服务器就可以实现.).
