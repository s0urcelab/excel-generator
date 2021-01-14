# excel-generator
> 🖨️ excel-generator是一款根据滴滴行程单快速生成报销用EXCEL的实用工具

## 快速开始
**推荐使用nodejs版本 v11+**

#### 在`app.js`中填写姓名和OA系统token
```javascript
const OA_TOKEN = '粘贴你OA系统cookie里的`j_authToken`字段到这里';

const USER_NAME = '这里填写你的姓名';
```
#### 行程单PDF放进 `/pdf` 文件夹

#### 执行下面命令，即可生成EXCEL
```javascript
npm i
npm run build
```
**生成的EXCEL将会被放置在`output`目录下**
