# <h1 align="center" >Jiangsu-Province nucleic-acid-information-batch-query<br/>江苏省核酸检测信息批量查询</h1>
<p align="center">
    <a href="https://github.com/W01fh4cker/nucleic-acid-information-query"><img alt="nucleic-acid-information-query" src="https://img.shields.io/github/stars/W01fh4cker/nucleic-acid-information-query.svg"></a>
    <a href="https://github.com/xzajyjs/ThunderSearch/releases"><img alt="nucleic-acid-information-query" src="https://img.shields.io/github/release/W01fh4cker/nucleic-acid-information-query.svg"></a>
    <a href="https://github.com/xzajyjs/ThunderSearch/issues"><img alt="nucleic-acid-information-query" src="https://img.shields.io/github/issues/W01fh4cker/nucleic-acid-information-query"></a>
    <a href="https://github.com/W01fh4cker/nucleic-acid-information-query"><img alt="nucleic-acid-information-query" src="https://img.shields.io/badge/python-3.7%20%7C%203.8%20%7C%203.9-blue"></a>
    <a href="https://github.com/W01fh4cker/nucleic-acid-information-query"><img alt="nucleic-acid-information-query" src="https://img.shields.io/github/followers/W01fh4cker?color=red&label=Followers"></a>
    <a href="https://github.com/W01fh4cker/nucleic-acid-information-query"><img alt="nucleic-acid-information-query" src="https://img.shields.io/badge/-%E6%A0%B8%E9%85%B8%E4%BF%A1%E6%81%AF%E6%9F%A5%E8%AF%A2-yellow"></a>
</p>  

## Introduction  
**调用江苏省核酸检测信息查询的`api`开发的一款江苏省核酸检测信息导入、查询、导出工具，适用于社区防疫人员核对工作，可以快速地把当地人员的身份信息导入，程序会自动查询核酸检测结果，返回结果并在当前文件夹生成结果导出的表格。  
有问题直接提issues，或者发信息至`sharecat2022@gmail.com`**  
## Attention  
### 1. 导入的表格姓名、身份证应分别放在第`1`、`3`列，如果不是，请修改源代码；  
### 2. 导入的表格名改为英文，且必须是`.xls`结尾，如果是`.xlsx`结尾，请打开后另存为`.xls`的；  
### 3. 导出的表格的空行代表需要手动查询。因为该用户有多个家庭成员，v1.1版本暂不支持查询家庭成员的核酸检测结果；  
## Version  
### v1.1 优化表格排版，优化匹配算法。  

### v1.0 实现表格导入匹配，并发送请求后接收内容，匹配后填入并导出表格。
