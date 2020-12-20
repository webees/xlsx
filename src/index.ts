import fs from 'fs'
import Excel from 'exceljs'

interface opt {
  src: string,
  dist: string,
  conf: string
}

export default class XLSX {
  static commissionNumber: string // 委托编号
  static commissionData: string // 委托日期
  static deviceNumber: string // 仪器编号
  static totalPages: number // 总页数
  static currentPage: number // 当前页
  static testName: string // 检测面
  static testAngle: string // 检测角度
  static testDate: string // 检测日期
  static pourfloor: string // 构件部位
  static pourName: string // 构件名称
  static pourDate: string // 浇筑日期
  static powerLevel: string // 强度等级
  static levelAge: string // 等级龄期
  static mpa1: string // 配置中的碳化深度
  static mpa2: string // 配置中的平均值
  static mpa3: string // 配置中的标准差
  static mpa4: string // 配置中的最小值
  static mpa5: string // 配置中的推定值
  static counter: number = 12 // 率定值计数器，每12次变更
  static fixValue = [] // 率定值

  static async start(opt: opt) {
    const files = (await fs.readdirSync(opt.src)).sort()
    const xlsx = await this.readXlsx(`${opt.src}/${opt.conf}`) // 读取配置
    this.commissionNumber = xlsx.getRow(1).values[1]
    this.commissionData = xlsx.getRow(2).values[8]
    this.deviceNumber = xlsx.getRow(7).values[8]
    this.totalPages = xlsx.lastRow.values[1]
    this.testName = xlsx.getRow(4).values[8]
    this.testAngle = xlsx.getRow(5).values[8]

    if ((files.length - 1) !== this.totalPages) throw '文件数与总页数不匹配！'
    console.log('总页数:', this.totalPages)
    console.log('委托编号:', this.commissionNumber)
    console.log('委托日期:', this.commissionData)
    console.log('仪器编号:', this.deviceNumber)
    console.log('检测面:', this.testName)
    console.log('检测角度:', this.testAngle)

    let sRow = 11 // 配置文件起始行 - 1
    for (let i = 0; i < files.length; i++) {
      if (files[i] === opt.conf) continue
      if (this.counter === 12) this.upFixValue(), this.counter = 0
      this.counter ++
      this.currentPage = i
      this.testDate = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[10]
      this.pourDate = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[4]
      this.pourfloor = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[3]
      this.pourName = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[2]
      this.powerLevel = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[11]
      this.levelAge = JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[12]
      this.mpa1 = parseFloat(JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[5]).toFixed(1)
      this.mpa2 = parseFloat(JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[6]).toFixed(1)
      this.mpa3 = parseFloat(JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[7]).toFixed(2)
      this.mpa4 = parseFloat(JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[8]).toFixed(1)
      this.mpa5 = parseFloat(JSON.parse(JSON.stringify(xlsx.getRow(sRow + i).values))[9]).toFixed(1)
      console.log(`\n\n文件名:${files[i]}`)
      console.log('检测日期:', this.testDate)
      console.log('浇筑日期:', this.pourDate)
      console.log('构件部位:', this.pourfloor)
      console.log('构件名称:', this.pourName)
      console.log('强度等级:', this.powerLevel)
      console.log('等级龄期:', this.levelAge)
      console.log('碳化深度:', this.mpa1)
      console.log('平均值:', this.mpa2)
      console.log('标准差:', this.mpa3)
      console.log('最小值:', this.mpa4)
      console.log('推定值:', this.mpa5)
      await this.step(`${opt.src}/${files[i]}`, `${opt.dist}/${files[i]}`)
    }
  }

  // 回弹法检测混凝土抗压强度原始记录
  static async step(src: string, dist: string) {
    // 输入文件
    const xlsx = await this.readXlsx(src)

    // 新文件
    const book = new Excel.Workbook()
    const sheet = book.addWorksheet('回弹法检测混凝土抗压强度原始记录')

    // 表头
    this.addHead(sheet)

    // 数据
    sheet.addRows(this.data(xlsx))

    // 合并
    this.mergeCells(sheet)

    // 样式
    this.setStyle(sheet)

    // 输出文件
    await book.xlsx.writeFile(dist);
  }

  // 原始值
  static data(sheet: Excel.Worksheet) {
    let rows = []
    // di 值位置随机
    const diRow = [4, 6, 8, 10, 12, 14, 16, 18]
    const di1Row = Math.floor(Math.random() * (3 - 0 + 1) + 0) // di1的行号
    const di2Row = Math.floor(Math.random() * (5 - (di1Row + 2) + 1) + (di1Row + 2)) // di2的行号
    const di3Row = Math.floor(Math.random() * (7 - (di2Row + 2) + 1) + (di2Row + 2)) // di3的行号
    let sRow = 3 // 开始行数
    let iRow = 3 // 间隔行数
    for (let i = sRow, j = 2; i < 33; i += iRow, j += 2) {
      if (j === diRow[di1Row]) rows = rows.concat(this.area(sheet, i, 1))
      else if (j === diRow[di2Row]) rows = rows.concat(this.area(sheet, i, 2))
      else if (j === diRow[di3Row]) rows = rows.concat(this.area(sheet, i, 3))
      else rows = rows.concat(this.area(sheet, i))
    }
    return rows
  }

  // 测区号
  static area(sheet: Excel.Worksheet, index: number, di: number = 0) {
    let rows = []
    const rm1Col = 16 // 回弹均值列号
    const rm2Col = 19 // 修正值列号
    const mpaCol = 20 // 测位推定
    const diRows = [0, 4, 7, 10] // 实测碳化深度行
    let area = JSON.parse(JSON.stringify(sheet.getRow(index).values)).slice(4, 25)
    let ri1 = area.slice(0, 8) // 回弹值列一
    let ri2 = area.slice(8, 16) // 回弹值列二
    let rm1 = area[rm1Col].toFixed(1) // 回弹均值
    let rm2 = area[rm2Col].toFixed(1) // 修正值
    let mpa = area[mpaCol].toFixed(1) // 测位推定
    ri1 = ri1.concat([rm1, rm2, mpa])
    ri2 = ri2.concat([null, null, null])
    if (di) { // 是否插入实测碳化深度
      let dis = this.di(sheet, diRows[di])
      ri1 = ri1.concat(dis.slice(0, 3))
      ri2 = ri2.concat(dis.slice(3, 4))
    }
    if (index === 3) {
      ri1 = ri1.concat([null, null, null])
      const other = this.other(sheet)
      ri1 = ri1.concat(other)
    }
    rows.push(ri1)
    rows.push(ri2)
    return rows
  }

  // 实测碳化深度 di
  static di(sheet: Excel.Worksheet, index: number) {
    let dis = JSON.parse(JSON.stringify(sheet.getRow(index).values)).slice(2, 5)
    for (let i = 0; i < 3; i++) {
      if (typeof dis[i] === 'string') {
        dis[i] = parseFloat(dis[i].replace(/\s*/g, '').substring(0, 4))
      }
      dis[i] = parseFloat(dis[i]).toFixed(2)
    }
    dis[3] = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(index + 1).values)).slice(3, 4)).toFixed(2)
    return dis
  }

  // 碳化深度、强度平均、标准差、最小值、强度推定、委托日期
  static other(sheet: Excel.Worksheet) {
    let rows = []
    const mpa1 = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(10).values))[4].replace(/\s*/g, '').substring(4)).toFixed(1)
    if (mpa1 !== this.mpa1) throw '碳化深度与总表不一致！'
    const mpa2 = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(33).values))[12].replace(/\s*/g, '').substr(-16, 4)).toFixed(1)
    if (mpa2 !== this.mpa2) throw '平均值与总表不一致！'
    const mpa3 = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(33).values))[12].replace(/\s*/g, '').substr(-12, 4)).toFixed(2)
    if (mpa3 !== this.mpa3) throw '标准差与总表不一致！'
    const mpa4 = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(33).values))[12].replace(/\s*/g, '').substr(-8, 4)).toFixed(1)
    if (mpa4 !== this.mpa4) throw '最小值与总表不一致！'
    const mpa5 = parseFloat(JSON.parse(JSON.stringify(sheet.getRow(33).values))[12].replace(/\s*/g, '').substr(-4, 4)).toFixed(1)
    if (mpa5 !== this.mpa5) throw '推定值与总表不一致！'
    rows.push(mpa1)
    rows.push(mpa2)
    rows.push(mpa3)
    rows.push(mpa4)
    rows.push(mpa5)
    const commissionData = this.commissionData.split('/')
    const testDate = this.testDate.split('/')
    const pourDate = this.pourDate.split('/')
    rows = rows.concat(commissionData)
    rows = rows.concat(testDate)
    rows = rows.concat(pourDate)
    rows.push(this.powerLevel)
    rows.push(this.levelAge)
    rows.push(this.deviceNumber)
    rows.push(this.testName)
    rows.push(this.testAngle)
    rows.push(this.pourfloor)
    rows.push(this.pourName)
    rows.push(this.currentPage)
    rows.push(this.totalPages)
    rows.push(this.commissionNumber)
    rows = rows.concat(this.fixValue)
    return rows
  }

  // 更新率定值
  static upFixValue() {
    let base = 80
    let fixValue = []
    for (let i = 0; i < 4; i++) {
      let tmp = 0
      for (; tmp === 0;) {
        const tmp1 = Math.floor(Math.random() * 2) // 0、1
        const tmp2 = Math.floor(Math.random() * 3) // 0、1、2
        if (tmp1) tmp = base + tmp2
        else tmp = base - tmp2
        for (let i = 0; i < fixValue.length; i++) if (fixValue[i] === tmp) tmp = 0
        for (let i = 0; i < this.fixValue.length; i++) if (this.fixValue[i] === fixValue[i] && fixValue[i] === tmp) tmp = 0
      }
      fixValue.push(tmp)
    }
    this.fixValue = fixValue
  }

  // 读取xlsx
  static async readXlsx(path: string, indexOrName: string | number = 1) {
    const book = new Excel.Workbook()
    const xlsx = await book.xlsx.readFile(path)
    return xlsx.getWorksheet(indexOrName); // 默认序数1
  }

  // 设置样式
  static setStyle(sheet: Excel.Worksheet) {
    sheet.eachRow(row => {
      row.eachCell(cell => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' }
      })
    })
  }

  // 添加表头
  static addHead(sheet: Excel.Worksheet) {
    sheet.mergeCells('A1:H1'), sheet.getCell('A1').value = '0'
    sheet.getCell('I1').value = '回弹均值'
    sheet.getCell('J1').value = '修正值'
    sheet.getCell('K1').value = '测位推定'
    sheet.mergeCells('L1:N1'), sheet.getCell('L1:N1').value = '实测碳化'
    sheet.getCell('O1').value = '碳化深度'
    sheet.getCell('P1').value = '平均值'
    sheet.getCell('Q1').value = '标准差'
    sheet.getCell('R1').value = '最小值'
    sheet.getCell('S1').value = '推定值'
    sheet.mergeCells('T1:V1'), sheet.getCell('T1').value = '委托日期'
    sheet.mergeCells('W1:Y1'), sheet.getCell('W1').value = '检测日期'
    sheet.mergeCells('Z1:AB1'), sheet.getCell('Z1').value = '浇筑日期'
    sheet.getCell('AC1').value = '强度等级'
    sheet.getCell('AD1').value = '等级龄期'
    sheet.getCell('AE1').value = '仪器编号'
    sheet.getCell('AF1').value = '检测面'
    sheet.getCell('AG1').value = '检测角度'
    sheet.getCell('AH1').value = '构件部位'
    sheet.getCell('AI1').value = '构件名称'
    sheet.getCell('AJ1').value = '当前页'
    sheet.getCell('AK1').value = '总页数'
    sheet.getCell('AL1').value = '委托编号'
    sheet.mergeCells('AM1:AP1'), sheet.getCell('AM1').value = '率定制'
  }

  // 合并单元格
  static mergeCells(sheet: Excel.Worksheet) {
    // I列合并
    for (let i = 2; i < 22; i += 2) {
      sheet.mergeCells(`I${i}:I${i + 1}`)
    }
    // J列合并
    for (let i = 2; i < 22; i += 2) {
      sheet.mergeCells(`J${i}:J${i + 1}`)
    }
    // K列合并
    for (let i = 2; i < 22; i += 2) {
      sheet.mergeCells(`K${i}:K${i + 1}`)
    }
    // LMN行合并
    for (let i = 3; i < 22; i += 2) {
      sheet.mergeCells(`L${i}:N${i}`)
    }
  }
}
