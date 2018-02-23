# 基于python的实验室自动化测试系统

或许是极人性化的一个文档管理框架，最适合部署在内网作为内网文档管理，url即目录层级。markdown+git+web搭配，让你一下子就喜欢上写文档分享。一分钟上手，有兴趣可挖掘隐藏技巧。

## 组网图

![walden](https://raw.github.com/meolu/Walden/master/static/screenshots/walden.gif)


## 特点

* Markdown语法
* 修改后实时展现，无编译
* 多模板支持
* 图片、附件上传，自动生成url
* 多项目
* 任意定义目录嵌套、定义文档，目录与文档均可中文（甚至推荐中文）
* 文档、图片、附件同步保存至git，这下你安心了吧

## 必备组件

### Python及相关库
* [python3.4.3](https://www.python.org/downloads/release/python-343/)
* [pyvisa 1.8](https://pypi.python.org/pypi/PyVISA/1.8)
* [pyserial 3.0](https://pypi.python.org/pypi/pyserial/3.0)
* [xlrd 1.1.0](https://pypi.python.org/pypi/xlrd/1.1.0)
* [xlwt 1.3.0](https://pypi.python.org/pypi/xlwt/1.3.0)
* [xlutils 2.0.0](https://pypi.python.org/pypi/xlutils/2.0.0)

### VISA
* [National Instruments’s VISA](http://www.ni.com/visa/)

### 安装说明：
考虑很多实验室仍然使用Windows XP系统，并且没有网络或者没有管理员权限，不能安装任何程序（作者公司提供的就是这种工作环境，吐槽一下）。
* Python选择的是最后一个支持xp版本3.4.3。pyserial选择的3.0。其他库没有特别要求，使用最新版即可。
* 可以在能够安装python和库的机器上，安装好全部库，打包整个python文件夹，直接拷贝到其他机器使用，在Sublime等文本编辑器中增加路径即可绿色化。

## 程序说明

```php
vi Config.php
return [
    // 项目留空保存文档和附件的git地址，可以是在github，好吧，不想公开，可以bitbucket。

    // 1.php进程的用户的id_rsa.pub已添加到git的ssh-key。这样才可以推送markdown下的文件。
    'git' => 'git@github.com:meolu/Walden-markdown-demo.git',

    // 2.好吧，如果实在不想加key，可以直接明文用户名密码认证的http(s)地址也可以。
    // 'git' => 'https://username:password@github.com/meolu/Walden-markdown-demo.git',
];
```

## 三、nginx简单配置

```
server {
    listen       80;
    server_name  Walden.dev;
    root /the/dir/of/Walden;
    index index.php;

    # 建议放内网做文档服务
    #allow 192.168.0.0/24;
    #deny all;

    location / {
        try_files $uri $uri/ /index.php$is_args$args;
    }

    location ~ \.php$ {
        try_files $uri =404;
        fastcgi_pass   127.0.0.1:9000;
        fastcgi_param  SCRIPT_FILENAME  $document_root$fastcgi_script_name;
        include        fastcgi_params;
    }
}
```


## 自定义模板

前端同学可以自己定义模板，在templates下新建一个模板目录，包含预览模板：`markdown-detail-view.php`，编辑模板：`markdown-editor-view.php`，然后修改`Config.php`的`template`为你的模板项目。

最后，当然希望你可以给此项目提个pull request，目前只有一个bootstrap的默认模板：(


## 下一步计划

* 修正：目前程序内部仍然有一处需要根据不同测试模块手工修改的缺陷。
* 修正：xlrd写入公式，必须手工打开excel并保存一次，才能被xlwt读取到计算值的缺陷。
* 找到合适的 1XN 光swtich，可以把被测端口全部连上，一次全部测完，无需手工拔插端口和修改代码选择不同被测端口。
* 优化代码，使程序更简洁优雅。
* 将所有配置修改部分，全部放入excel表格，做到程序完全黑盒化。

## 版本记录

版本记录：[CHANGELOG](https://github.com/yugiyx/happy-send/blob/master/CHANGELOG.md)




