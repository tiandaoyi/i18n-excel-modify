#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 获取命令行参数，第一个参数是 Node.js 的执行路径，第二个参数是脚本文件路径，之后的参数是传递的参数
const fileName = process.argv[2];

if (!fileName) {
  console.error('请提供 Excel 文件名：node main.js xx.excel');
  process.exit(1);
}

const filePath = path.join(process.cwd(), fileName);

// 读取 Excel 文件
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0]; // 假设只有一个 Sheet

// 获取 Sheet 中的数据
const sheet = workbook.Sheets[sheetName];

// 找到表头为 '英文' 的列
let targetColumn;
for (let col in sheet) {
  // if (col.startsWith('A') && sheet[col].v === '英文') {
  if (sheet[col].v === '英文') {
    targetColumn = col[0];
    break;
  }
}

if (!targetColumn) {
  console.error('未找到表头为 "英文" 的列');
  process.exit(1);
}

// 遍历该列的每个单元格
for (let row in sheet) {
  if (row.startsWith(targetColumn) && row !== targetColumn + '1') {
    const cellValue = sheet[row].v;
    if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
      // 执行修改方法
      sheet[row].v = classifyString(cellValue);
    }
  }
}

const extname = path.extname(fileName);
const backupFileName = `${fileName.replace(extname, `-${Date.now()}-bak${extname}`)}`;

const backupFilePath = path.join(process.cwd(), backupFileName);
// 备份
fs.copyFileSync(filePath, backupFilePath);

// 保存修改后的文件
XLSX.writeFile(workbook, filePath);

console.log(`文件处理完成: ${fileName}`);
console.log(`源文件备份完成: ${backupFileName}`);

function formatString(input) {
  // 使用正则表达式将每个单词的首字母大写
  return input.replace(/\b\w/g, function (match) {
    return match.toUpperCase()
  })
}

function classifyString(input) {
  // 使用正则表达式检查输入是否为单词
  const wordPattern = /^[a-zA-Z]+$/

  const phrasePattern = /[.,;!]/

  if (wordPattern.test(input)) {
    return formatString(input)
  } else if (!phrasePattern.test(input)) {
    const words = input.split(/\s+/)
    if (words.length >= 4) {
      return input.charAt(0).toUpperCase() + input.slice(1)
    } else {
      return formatString(input)
    }
  } else {
    return input.charAt(0).toUpperCase() + input.slice(1)
  }
}
