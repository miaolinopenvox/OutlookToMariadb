﻿# OutlookToMariadb

Export emails from MS Outlook to mysql/mariadb periodically.  
定期将MSOutlook 中的电子邮件导入mysql/mariadb.  

## Development Tool(开发工具):  
Visual Studio 2022  
c#  
.net core 6.0LTS  

## Files in this project(项目文件):  
- outlook.sql: used to create database and tables, views in mariadb. 用于在mysql中创建数据库和表，view.  
- OutlookToMariadb.ini: config file for this tool. 配置文件  
    - See comments in this file.
    - 内容涵义参考配置文件内的注释  

## Database Tables(数据库表):  
### email: 
- Email content and with allmost all fields from MSOutlook.   
- 来自MSOutlook的email 内容和大多数必要的字段  
### folders:  
- All the FullPathName of folders in MSOutlook.  
- MSOutlook中所有的子文件夹全路径名。
### fulltext_email: 
- A view of table 'email' that bombine all the human readable text fields in to 'fulltexts' fields.  
- 来自表email的视图，将所有表email的文本字段合并成为'fulltexts'字段，方便全文检索。   


## Field 'action' in table 'email'(表email中的action字段):  
OutlookToEmail will check this field periodically to execute operate to this mail item.  
OutlookToEmail 周期行的检查这个字段，并根据字段内容对这一条mail执行相应的操作.  

### Right now following string as operations was supported(目前支持以下字符串作为操作):
- "delete"
    - Delete this mail from MS Outlook.  
    - 从MS Outlook删除此邮件.  
- "markreaded"
    - Mark this mail as readed in  MS Outlook.  
    - 在MS Outlook中将此邮件标记为已读.  

## How to use(如何使用):
- Use SQL manipulate emails out of outlook(使用sql语句在outlook外操作邮件)  
- Write you own anti-spam tool with in programming language(使用任何编程语言编写垃圾邮件过滤器).