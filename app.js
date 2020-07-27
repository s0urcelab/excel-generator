/**
 *  @Author  : s0urce <apao@douyu.tv>
 *  @Date    : 2020/6/29
 *  @Declare : generate EXCEL
 *
 */
const CONFIG = {
    name: '小明',
    OAToken: 'ZaaaaaX',
    pdfPath: './assets/滴滴出行行程报销单.pdf',
    tplPath: ['./assets/c.xlsx', './assets/d.xlsx'],
    outputPath: ['./dist/报销申请单.xlsx', './dist/市内交通费用报销明细.xlsx'],
}

const moment = require('dayjs')
const pdfjs = require('pdfjs-dist/es5/build/pdf')
const Excel = require('exceljs')
const rmb = require('rmb-x')
const axios = require('axios')
const NP = require('number-precision')

const DATE_FORMAT = 'YYYY/M/D'
const STAT_REG = /共(\d+)笔行程， 合计 ((\d|\.)+)元/
const START_ROW_IDX_C = 5
const START_ROW_IDX_D = 4
const TODAY = moment().format(DATE_FORMAT)
const THIS_YEAR = moment().format('YYYY')

const REQ_OPTS = {
    url: 'https://dymhr.douyucdn.cn/kq/japi/attend/record/list',
    method: 'POST',
    headers: {
        Cookie: `j_authToken=${CONFIG.OAToken}`,
        Host: 'dymhr.douyucdn.cn',
        Origin: 'https://dymhr.douyucdn.cn',
        Referer: 'https://dymhr.douyucdn.cn/',
    },
    data: {
        attendStartDate: moment()
            .startOf('month')
            .hour(0)
            .minute(0)
            .second(0)
            .millisecond(0)
            .unix(),
        attendEndDate: moment()
            .endOf('month')
            .hour(0)
            .minute(0)
            .second(0)
            .millisecond(0)
            .unix(),
        page: 1,
        pageSize: 30,
        unum: '201909212',
    },
}

const workbook = new Excel.Workbook();

(async () => {
    console.log('开始解析行程单...')
    try {
        const loadingPdf = pdfjs.getDocument(CONFIG.pdfPath)
        const doc = await loadingPdf.promise

        const firstPage = await doc.getPage(1)
        const { items } = await firstPage.getTextContent()
        const [_, statNum, statMoney] = items
            .find(v => !!~v.str.indexOf('笔行程， 合计')).str
            .match(STAT_REG)

        let tableList = []
        for (let i = 1; i <= doc.numPages ; i++) {
            const nowPage = await doc.getPage(i)
            const { items } = await nowPage.getTextContent()
            const tableData = items.filter(v => String(v.transform.slice(0, 4)) === '7,0,0,7')
            const groupCount = Math.ceil(tableData.length / 10)
            const pageList = Array.from({length: groupCount}).map((item, idx) => {
                const [date, time, week] = tableData[idx * 10 + 2].str.split(' ')
                return {
                    type: tableData[idx * 10 + 1].str,
                    date: moment(`${THIS_YEAR}-${date}`).format(DATE_FORMAT),
                    time,
                    startLocation: tableData[idx * 10 + 4].str,
                    endLocation: tableData[idx * 10 + 5].str,
                    distance: +tableData[idx * 10 + 6].str,
                    money: +tableData[idx * 10 + 7].str,
                }
            })
            tableList = [...tableList, ...pageList]
        }

        const totalNum = tableList.length
        const totalMoney = tableList.reduce((acc, curr) => {
            return NP.plus(acc, curr.money)
        }, 0)

        if (totalNum !== +statNum) {
            return console.error(`行程总笔数不一致：${totalNum}, ${statNum}`)
        }
        if (totalMoney !== +statMoney) {
            return console.error(`合计金额不一致：${totalMoney}, ${statMoney}`)
        }

        // 报销申请单
        await workbook.xlsx.readFile(CONFIG.tplPath[0])

        const sheetC = workbook.getWorksheet('报销单')
        sheetC.getCell('A18').value = `日期： ${TODAY}`
        sheetC.getCell('F6').value = totalMoney
        sheetC.getCell('K11').value = totalMoney
        sheetC.getCell('K14').value = rmb(totalMoney)
        // copy
        sheetC.duplicateRow(START_ROW_IDX_C, totalNum - 1, true)
        tableList.forEach((item, idx) => {
            sheetC.getCell(`A${idx + START_ROW_IDX_C}`).value = idx + 1
            sheetC.getCell(`B${idx + START_ROW_IDX_C}`).value = item.date
            sheetC.getCell(`F${idx + START_ROW_IDX_C}`).value = item.money
        })

        // output
        await workbook.xlsx.writeFile(CONFIG.outputPath[0])
        console.log('output:', CONFIG.outputPath[0])

        // 市内交通费用报销明细
        const response = await axios(REQ_OPTS)
        const { data: { data: { records } } } = response
        const dList = tableList.map(item => {
            const {endCheckTime} = records
                .find(v => moment.unix(v.attendDate).format(DATE_FORMAT) === item.date)
            return {
                ...item,
                endCheck: moment.unix(endCheckTime).format('HH:mm'),
            }
        })

        await workbook.xlsx.readFile(CONFIG.tplPath[1])
        const sheetD = workbook.getWorksheet('Sheet1')
        sheetD.getCell('I2').value = TODAY
        sheetD.getCell('H5').value = totalMoney
        // copy
        sheetD.duplicateRow(START_ROW_IDX_D, totalNum - 1, true)
        dList.forEach((item, idx) => {
            sheetD.getCell(`A${idx + START_ROW_IDX_D}`).value = item.date
            sheetD.getCell(`B${idx + START_ROW_IDX_D}`).value = item.endCheck
            sheetD.getCell(`C${idx + START_ROW_IDX_D}`).value = CONFIG.name
            sheetD.getCell(`D${idx + START_ROW_IDX_D}`).value = item.startLocation
            sheetD.getCell(`E${idx + START_ROW_IDX_D}`).value = item.endLocation
            sheetD.getCell(`G${idx + START_ROW_IDX_D}`).value = item.money
            sheetD.getCell(`H${idx + START_ROW_IDX_D}`).value = item.money
            sheetD.getCell(`J${idx + START_ROW_IDX_D}`).value = CONFIG.name
        })

        // output
        await workbook.xlsx.writeFile(CONFIG.outputPath[1])
        console.log('output:', CONFIG.outputPath[1])
    } catch (e) {
        console.error('throw error:', e)
    }
})()
