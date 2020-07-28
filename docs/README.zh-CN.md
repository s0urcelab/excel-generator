# excel-generator
> 🖨️ excel-generator是一款根据滴滴行程单快速生成报销用EXCEL的实用工具

## 快速开始
**推荐使用nodejs版本 v11+**

在`app.js`中填写姓名和OA系统token
```javascript
const CONFIG = {
    name: '小明', // 填写你的姓名
    OAToken: 'ZaaaaaX', // 粘贴你OA系统cookie里的`j_authToken`字段
    pdfPath: './assets/滴滴出行行程报销单.pdf', // 将滴滴行程单pdf放置在此路径下
    tplPath: ['./assets/c.xlsx', './assets/d.xlsx'], // default template EXCEL
    outputPath: ['./dist/报销申请单.xlsx', './dist/市内交通费用报销明细.xlsx'], // output path setting
}
...
```
执行下面命令，即可生成EXCEL
```javascript
npm i
npm run build
```
**生成的EXCEL将会被放置在`dist`目录下**
