# excel-generator
🖨️ a tool to generate EXCEL

## Usage
Edit `app.js`
```javascript
const CONFIG = {
    name: '小明', // replace your real name
    OAToken: 'ZaaaaaX', // paste your OA j_authToken here
    pdfPath: './assets/滴滴出行行程报销单.pdf', // put your pdf here
    tplPath: ['./assets/c.xlsx', './assets/d.xlsx'], // default template EXCEL
    outputPath: ['./dist/报销申请单.xlsx', './dist/市内交通费用报销明细.xlsx'], // output path setting
}
...
```
Build EXCEL
```javascript
npm i
npm run build
```

## Develop
```javascript
npm i
npm start
```
