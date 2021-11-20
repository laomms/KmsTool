# KMS Local Activation Tool  

#### support win7/8.1/10/server2022 office2010~office2021

![image](https://github.com/laomms/KmsTool/blob/main/kms.JPG)     
![image](https://github.com/laomms/KmsTool/blob/main/kms2.png)   

分流网盘: https://bkkmms.herokuapp.com/BAIDUD/Activation   

##### 软件工作原理:  
KMS本地激活:  
KMS激活逃不过KMS连接代理:SppExtComObj.Exe这个软件,这个工具在系统目录,验证KMS和合法性全部通过这个,只要对这个可执行文件中的RpcStringBindingComposeW和RpcBindingFromStringBindingW进行拦截,让他将KMS服务器指向本地就可以,这样可以在注册表中搭建一个KMS服务器,实现本地激活(早期Windows版本没有这个工具,直接在注册版模拟服务器就可以实现.).

数字激活:  
通过调用系统自带Clipup.exe(或者镜像中的gatherosstate.exe)中的HwidGetCurrentEx获取本机唯一硬件ID,结合系统的相应版本的PFN生成SessionId,然后通过GetGenuineTicket函数生成数字证书.通过安装密钥向微软服务器获取用户证书,跟安装的数字证书进行对比验证激活.

SLIC注入激活:  
通过GetVolumePathNamesForVolumeNameW打开启动分区,加入某品牌的SLIC表信息,然后安装相应的OEM:SLIC密钥,通过对比SLIC表中的公钥来验证激活.支持MBR和GPT分区格式.

密钥激活:  
通过SLInstallProofOfPurchase安装密钥,通过SLGetPKeyInformation获取密钥的SKUID,通过SLActivateProduct激活系统.


这个工具包括5个激活模块：

1、数字权利激活，适用于Win10/Win11 RTM版本及LTSB2015-LTSC2021的激活。通过硬件生成的唯一有效数字证书连接微软服务器获取永久激活。  
2、KMS38激活，适用于Win10/Win11系统全系列，通过硬件生成的唯一有效数字证书连接微软服务器获取KMS激活，激活至2038年。  
3、KMS本机激活，适合Office2010-Office2021激活(也适用于windows激活)。原理是在注册表中模拟KMS服务器,通过对SppExtComObj进行拦截指向本地服务器实现KMS180天激活.可设置成到期无限循环。变相达到永久激活目的。  
4、Slic注入激活，适用于Win7全系列(除企业版外)及Server2008-Server2022常用版本激活.通过在启动分区中注入品牌电脑SLIC标识信息实现OEM离线激活。. 
5、密钥激活，安装有效密钥并通过自动获取确认ID实现激活。   
  
所以在没有密钥的情况下，该工具基本上可以实现VISTA版本后的全部激活。















