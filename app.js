var Excel = require('exceljs');
const mock = require('./mock')

const separate_way = (text) => {
    if (mock.กทม.filter(x => x == text).length > 0)
        return 'กทม ปริมณฑล'
    else if (mock.ภาคกลาง.filter(x => x == text).length > 0)
        return 'ภาคกลาง'
    else if (mock.ภาคเหนือ.filter(x => x == text).length > 0)
        return 'ภาคเหนือ'
    else if (mock.ภาคตะวันออกเฉียงเหนือ.filter(x => x == text).length > 0)
        return 'ภาคตะวันออกเฉียงเหนือ'
    else if (mock.ภาคตะวันออก.filter(x => x == text).length > 0)
        return 'ภาคตะวันออก'
    else if (mock.ภาคตะวันตก.filter(x => x == text).length > 0)
        return 'ภาคตะวันตก'
    else if (mock.ภาคใต้.filter(x => x == text).length > 0)
        return 'ภาคใต้'
    else
        return '-'
}

const main = async () => {
    try {
        var Workbook = new Excel.Workbook()
        await Workbook.xlsx.readFile('test.xlsx')
        // var Worksheet = Workbook.getWorksheet(7)
        await Workbook.eachSheet(async (Worksheet, sheetId) => {
            if (sheetId != 2) {
                await Worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    if (rowNumber > 3) {
                        const d = row.values[10]
                        Worksheet.getCell(`P${rowNumber}`).value = separate_way(d)
                        console.log(rowNumber)
                    }
                })
            } else if(sheetId == 2) {
                await Worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    if (rowNumber > 3) {
                        const d = row.values[9]
                        Worksheet.getCell(`O${rowNumber}`).value = separate_way(d)
                        console.log(rowNumber)
                    }
                })
            }
        })

        await Workbook.xlsx.writeFile('new.xlsx');
    } catch (error) {
        console.error('Error editing Excel file:', error);
    }
}

main()