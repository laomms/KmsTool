##### 本工具只可用来做个人系统证书环境调试用，一切与之有关的商业活动都是非法的。请支持正版软件。

# KMS Local Activation Tool  
基于Windows证书的一个激活工具，支持Win10/11 ARM64。  
优势:
- 可选性激活.可单独激活某项而不影响其他证书。  
- 支持生成几乎所有win10/11版本的数字证书(分零售和批量).相关密钥提取自win10最新版本22000,该支持的都支持.   
- 为win10/11以下版本提供完美解决方案.    
- 为Office系列提供本地KMS解决方案.   

##### 数字权利 [支持版本](README_HWID.md) 
适用于大部分Windows10/11RTM版本。基于Windows硬件指纹的一种激活方式。分零售和批量方式(KMS38). 

##### KMS本地激活 [支持版本](README_KMS.md)
无需远程KMS服务器，直接通过HOOK SppExtComObj.exe及在注册表中模拟KMS服务器信息获取180天的批量授权许可，可以循环使用。内置几乎所有版本的KMS密钥。

##### SLIC注入激活 [支持版本](README_OEMSLIP.md)
基于OEM SLIC密钥的一种离线激活方式，只要有OEM:SLP密钥的系统基本上都激活。通过模拟某品牌的SLIC表信息,然后安装相应的OEM:SLIC密钥,通过对比SLIC表中的公钥来验证激活.支持MBR和GPT分区格式.适用于win7系列及server系列的激活。

![image](https://github.com/laomms/KmsTool/blob/main/kms.JPG)     
![image](https://github.com/laomms/KmsTool/blob/main/kms2.png)   



















