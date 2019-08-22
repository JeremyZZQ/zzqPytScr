# 记录在vscode中设置Git的操作 #

## 安装软件 ##

注意，在Windows上安装本应很容易，但因为我参照cyggit的教程使用cyggit安装git，导致vscode配置怎么弄都不生效（包括git-path），后来直接装git（默认目录），然后修改git-path配置，重启就好了，不会再提示代码管理程序没有注册。

### 安装Git后的设置 ###

git config --global user.name "JeremyZZQ" 
git config --global user.email "zhangzhiqiang81@yeah.net"

## SSH通信设置 ##

### github的设置 ###

#### 在git本地生成SSHKey ####
ssh-keygen -t rsa -C "zhangzhiqiang81@yeah.net"

需要三次确认（回车，如果是默认）
1、设置SSHKey保存文件，默认是C盘用户文件夹下的.ssh文件夹里的id_rsa.pub文件。
2、输入密码
3、确认密码

#### 在Github设置 ####
1、创建新的项目和库。
2、在设置（setting >> SSH and GPG keys）中添加上述id_rsa.pub文件中记载的密文全文。

#### 在Git Bash中测试通讯 ####

ssh -T git@github.com

## 在Git Bash中克隆Github上的仓库 ##

1、首先从github上复制仓库的地址（ssh的）
2、在git bash中克隆：
git clone git@github.com:地址

这样，在vscode的Git中就能添加仓库或看到仓库的内容