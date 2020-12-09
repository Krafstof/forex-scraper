const puppeteer = require('puppeteer') //библиотека для headless браузера
var excel = require('exceljs'); //библиотека для взаимодействия с excel и .csv файлами
const csv = require('csv-parser')//библиотека для взаимодействия с csv
const fs = require('fs');//файловый модуль javascript

let csvData = []

fs.createReadStream('EURUSD1440.csv')
    .pipe(csv())
    .on('data', (data) => csvData.push(data))//запись данных из CSV файла в массив csvData
    .on('end', () => {
        console.log('CSV file reading complete');
        csvData.forEach((element, index) => {
            element.Date = element.Date.toString().substring(8, 10) +
                '.' + element.Date.toString().substring(5, 7) +
                '.' + element.Date.toString().substring(0, 4)//преобразование времени 
        })
    })

scrapeForexNews()

async function scrapeForexNews() {
    const browser = await puppeteer.launch({ headless: true })//запуск браузера

    //создание excel файла вывода
    var workbook1 = new excel.Workbook();
    workbook1.creator = 'Me';
    workbook1.lastModifiedBy = 'Me';
    workbook1.created = new Date();
    workbook1.modified = new Date();
    console.log(workbook1);
    var sheet1 = workbook1.addWorksheet('Sheet1');
    var reColumns = [ //определение заголовков
        { header: 'Дата', key: 'a' },
        { header: 'Дата прогноза', key: 'z' },
        { header: 'Прогноз', key: 'k' },
        { header: 'Открытие', key: 'l' },
        { header: 'Закрытие', key: 'c' },
        { header: 'Максимум', key: 'd' },
        { header: 'Минимум', key: 'e' },
        { header: 'Объем', key: 'h' },
        { header: 'Верен ли бы прогноз', key: 'o' },
        { header: 'Ссылка на прогноз', key: 'f' },
    ];
    sheet1.columns = reColumns;
    let goodPredictions = 0//количество сбывшихся прогнозы
    let allPredictions = 0//количество всех прогнозов
    const page = await browser.newPage()
    for (let i = 17; i <= 167; i++) {//количество итераций в цикле равно количеству страниц со всеми прогнозами по паре EUR USD
        await page.goto(`https://ru.forexnews.pro/category/prognozy_forex/%d0%bf%d1%80%d0%be%d0%b3%d0%bd%d0%be%d0%b7-eurusd/page/${i}/`) //ссылка на прогнозы с сайта ForexNews
        const itemRaw = await page.evaluate(
            () => Array.from(document.querySelectorAll('.post.type-post a')).map((weblink) => weblink.href) //считывание ссылок с прогнозами
        )
        const itemFiltered = itemRaw.filter(link => link.toString().includes('eurusd-prognoz-evro-dollar')
            && link.toString().includes('respond')
            && !link.toString().includes('nedelyu')) //фильтрация ссылок из-за особенностей их описания на сайте
        for (let j = 0; j < itemFiltered.length; j++) {
            const subPage = await browser.newPage()
            await subPage.goto(itemFiltered[j]) //переход на ссылку с прогнозом

            let dateOfPrediction = await subPage.evaluate(
                () => Array.from(document.querySelectorAll('.published')).map((date) => date.innerText) //считывание даты прогноза (всегда предыдущий день перед тем, когда ожидается прогнозируемое значение)
            )
            dateOfPrediction = dateOfPrediction.toString()

            const allParagraphs = await subPage.evaluate(
                () => Array.from(document.querySelectorAll('.entry-content')).map((para) => para.innerText) //считывание всего текста статьи
            )
            let start = allParagraphs.toString().lastIndexOf('1,')
            let endSpace = allParagraphs.toString().indexOf(' ', start)
            let endDot = allParagraphs.toString().indexOf('.', start)
            let prediction = '-1'
            if (endSpace < endDot) prediction = allParagraphs.toString().substring(start, endSpace) //определение котировки прогноза
            else prediction = allParagraphs.toString().substring(start, endDot)
            if (prediction[prediction.length - 1] === ',') prediction[prediction.length - 1] = ''
            prediction = prediction.replace(',', '.') //замена запятой в прогнозе на точку

            let date = 'test'
            let opening = 'test'
            let closure = 'test'
            let maximum = 'test'
            let minimum = 'test'
            let volume = 'test'
            let predictionResult = 'test'
            for (let k = 0; k < csvData.length - 1; k++) {
                if (csvData[k].Date == dateOfPrediction) {
                    date = csvData[k + 1].Date
                    closure = csvData[k + 1].Closure
                    opening = csvData[k + 1].Open
                    maximum = csvData[k + 1].Max
                    minimum = csvData[k + 1].Min
                    volume = csvData[k + 1].Volume
                    let closurePredictionDay = csvData[k].Date
                    closurePredictionDay = parseFloat(closurePredictionDay)
                    allPredictions++
                    let predictionFloat = parseFloat(prediction)
                    let maximumFloat = parseFloat(maximum)
                    let minimumFloat = parseFloat(minimum)
                    let openingFloat = parseFloat(opening)
                    if (predictionFloat > closurePredictionDay) {
                        if (predictionFloat <= maximumFloat) {
                            predictionResult = 'да, повышение'
                            goodPredictions++
                        }
                        else predictionResult = 'нет, повышение'
                    }
                    else if (predictionFloat < closurePredictionDay) {
                        if (predictionFloat >= minimumFloat) {
                            predictionResult = 'да, понижение'
                            goodPredictions++
                        }
                        else predictionResult = 'нет, понижение'
                    }
                }
            }

            const rows = [
                [date, dateOfPrediction, prediction, opening, closure, maximum, minimum, volume, predictionResult, itemFiltered[j]]//создание ряда данных в итоговой таблице
            ]
            sheet1.addRows(rows);//запись данных
            console.log(rows);
            await subPage.close()
        }
    }
    const rating = goodPredictions / allPredictions
    console.log(rating);
    const lastRow = [
        [' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', 'Рейтинг аналитика = ' + rating, '']//создание ряда финальных данных в итоговой таблице
    ]
    sheet1.addRows(lastRow)
    workbook1.xlsx.writeFile("./uploads/error.xlsx").then(function () { //запись и закрытие итоговой таблицы
        console.log("xlsx file is written.");
    });
    await browser.close()
}
