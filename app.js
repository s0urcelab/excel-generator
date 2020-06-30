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

const moment = require('moment')
const pdfjs = require('pdfjs-dist/es5/build/pdf')
const Excel = require('exceljs')
const rmb = require('rmb-x')
const request = require('request');

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
        cookie: `j_authToken=${CONFIG.OAToken}`,
        host: 'dymhr.douyucdn.cn',
        origin: 'https://dymhr.douyucdn.cn',
        referer: 'https://dymhr.douyucdn.cn/',
    },
    body: {
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
        unum: "201909212",
    },
    json: true,
};

const workbook = new Excel.Workbook()
const loadingPdf = pdfjs.getDocument(CONFIG.pdfPath)
loadingPdf.promise
    .then(function (doc) {
        doc.getPage(1).then(function (page) {
            return page
                .getTextContent()
                .then(function (content) {
                    const {items} = content
                    const [_, statNum, statMoney] = items
                        .find(v => !!~v.str.indexOf('笔行程， 合计')).str
                        .match(STAT_REG)
                    const tableData = items.filter(v => String(v.transform.slice(0, 4)) === '7,0,0,7')
                    const groupCount = Math.ceil(tableData.length / 10)
                    const tableList = Array.from({length: groupCount}).map((item, idx) => {
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

                    const totalNum = tableList.length
                    const totalMoney = tableList.reduce((acc, curr) => {
                        return acc + curr.money
                    }, 0)

                    if (totalNum !== +statNum) {
                        throw `行程总笔数不一致：${totalNum}, ${statNum}`
                    }
                    if (totalMoney !== +statMoney) {
                        throw `合计金额不一致：${totalMoney}, ${statMoney}`
                    }

                    return {
                        list: tableList,
                        totalNum,
                        totalMoney,
                    }
                })
                .then(({list, totalNum, totalMoney}) => {
                    // 报销申请单
                    workbook.xlsx.readFile(CONFIG.tplPath[0])
                        .then(() => {
                            const worksheet = workbook.getWorksheet('报销单')
                            worksheet.getCell('A18').value = `日期： ${TODAY}`
                            worksheet.getCell('F6').value = totalMoney
                            worksheet.getCell('K11').value = totalMoney
                            worksheet.getCell('K14').value = rmb(totalMoney)
                            // copy
                            worksheet.duplicateRow(START_ROW_IDX_C, totalNum - 1, true)
                            list.forEach((item, idx) => {
                                worksheet.getCell(`A${idx + START_ROW_IDX_C}`).value = idx + 1
                                worksheet.getCell(`B${idx + START_ROW_IDX_C}`).value = item.date
                                worksheet.getCell(`F${idx + START_ROW_IDX_C}`).value = item.money
                            })

                            // output
                            workbook.xlsx.writeFile(CONFIG.outputPath[0])
                                .then(() => {
                                    console.log('output:', CONFIG.outputPath[0])
                                })

                        })

                    // 市内交通费用报销明细
                    request(REQ_OPTS, function (error, response, body) {
                        if (error) throw error
                        const { data: { records = [] } = {} } = body
                        const dList = list.map(item => {
                            const { endCheckTime } = records
                                .find(v => moment.unix(v.attendDate).format(DATE_FORMAT) === item.date)
                            return {
                                ...item,
                                endCheck: moment.unix(endCheckTime).format('HH:mm'),
                            }
                        })

                        workbook.xlsx.readFile(CONFIG.tplPath[1])
                            .then(() => {
                                const worksheet = workbook.getWorksheet('Sheet1')
                                worksheet.getCell('I2').value = TODAY
                                worksheet.getCell('H5').value = totalMoney
                                // copy
                                worksheet.duplicateRow(START_ROW_IDX_D, totalNum - 1, true)
                                dList.forEach((item, idx) => {
                                    worksheet.getCell(`A${idx + START_ROW_IDX_D}`).value = item.date
                                    worksheet.getCell(`B${idx + START_ROW_IDX_D}`).value = item.endCheck
                                    worksheet.getCell(`C${idx + START_ROW_IDX_D}`).value = CONFIG.name
                                    worksheet.getCell(`D${idx + START_ROW_IDX_D}`).value = item.startLocation
                                    worksheet.getCell(`E${idx + START_ROW_IDX_D}`).value = item.endLocation
                                    worksheet.getCell(`G${idx + START_ROW_IDX_D}`).value = item.money
                                    worksheet.getCell(`H${idx + START_ROW_IDX_D}`).value = item.money
                                    worksheet.getCell(`J${idx + START_ROW_IDX_D}`).value = CONFIG.name
                                })

                                // output
                                workbook.xlsx.writeFile(CONFIG.outputPath[1])
                                    .then(() => {
                                        console.log('output:', CONFIG.outputPath[1])
                                    })
                            })
                    });
                })
                .catch(console.error)
        })
    })




