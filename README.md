# i18n-excel-modify

该脚本可以将excel中的“英文”表头的列统一处理大小写

## 转换大小写的规则

1. 首字母统一大写
2. 如果是句子，则只有第一个单词首字母大写
3. 如果非句子，每个单词首字母大写

句子判断方法，包含逗号或者句号，或者大于等于5个以上单词。

## 使用

npm run start 【文件名】.xlsx
