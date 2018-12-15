# Excel生成Android Stings.xml工具

> fork from [liliLearn/generate_android_strings_tool](https://github.com/liliLearn/generate_android_strings_tool)
>
> 原来的工程只能单个的strings.xml和Excel互相转换，在实际工作中这样子仍然不方便。
>
> 在此基础上改动了一下：
>
> **1/2** 可以自动导入`res`目录下所有`values**/strings.xml`或者`values**/string.xml`文件，并输出到同一个Excel文件中。
>
> **2/2** 同理，一份这样的Excel文件也可以转化为多个`values**/strings.xml`文件，使用起来更加方便。
>
> 可以直接下载使用 [点我下载⬇️](https://github.com/jixiaoyong/generate_android_strings_tool/raw/master/release/generate_android_strings_tool_20181215.jar)
>
> 感谢原作者 **[liliLearn](https://github.com/liliLearn)**

------

![WX20180427-134909.png](https://upload-images.jianshu.io/upload_images/5488544-e1594caa69c3184b.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![WX20180427-134918.png](https://upload-images.jianshu.io/upload_images/5488544-c5ce4feebbce4ff8.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![选择res目录](http://jixiaoyong.github.io/images/20181215200417.png)


#### Excel 表格格式
key | cn | en |ja
---|---|---|---
login |	登录	| Login	 | 登録
name		| 姓名	| 	name		| ユーザーネーム
mail_address		| 邮箱	| 	Mail address		| メールアドレス
password	| 	密码		| Password		| パスワード

#### 支持注释

![WX20180427-150935.png](https://upload-images.jianshu.io/upload_images/5488544-f19fafad39a0b45b.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

#### 生成结果
```xml
<?xml version="1.0" encoding="UTF-8"?>
<resources>
  <!--test-->
  <string name="login">登录</string>
  <string name="name">姓名</string>
  <string name="mail_address">邮箱</string>
  <string name="password">密码</string>
</resources>
```

**生成的Excel**

![生成的XML文件](http://jixiaoyong.github.io/images/20181215200746.png)

**生成的XML**

![生成的XML文件](http://jixiaoyong.github.io/images/20181215200825.png)

#### *注意事项

    key：固定标识
    支持注释：key列可以使用注释（直接在Excel中写入注释）
    完善表格 别出现空行，不会报错但是会写空字符串
