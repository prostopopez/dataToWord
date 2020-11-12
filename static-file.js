// Express, tempfile, officegen
let express = require('express');
let path = require('path');
let tempfile = require('tempfile');
let officegen = require('officegen');
let docx = officegen({
    type: 'docx',
    subject: 'Report Name',
    keywords: 'Reports',
    orientation: 'portrait',
    description: 'A Report Description',
    pageMargins: { top: 100, left: 400, bottom: 100, right: 400 }
});

// Express-приложение
let app = express();
// Подключение статичных файлов
app.use("/static", express.static('./static/'));
// Маршрут для главной страницы
app.get('/', function (req, res) {
    res.sendFile('index.html', { root: __dirname });
});

// Парсер
const parseData = (req, res, next) => {
    if (req.method === 'POST') {
        const formData = {};
        req.on('data', data => {

            // Расшифровка данных
            const parsedData =
                decodeURIComponent(data).split('&');

            for (let data of parsedData) {
                decodedData = decodeURIComponent(
                    data.replace(/\+/g, '%20'));

                const [key, value] =
                    decodedData.split('=');

                // Данные => объект
                formData[key] = value;
            }

            // Данные формы => объект запроса
            req.body = formData;
            next();
        });
    } else {
        next();
    }
};

// Вывод информации
app.post('/finalData', parseData, (req, res) => {
    // Получение данных формы из объекта запроса
    const data = req.body;
    let {
        mainDate, cardNumber, paramedicOnCall,
        callReceiveTime, callTransferTime, teamCallTime, arrivalTime, transportStartTime, transportEndTime, callEndTime, isAirTransport,
        startCity, startHealthCareFacility, startPhoneNumber, startNameAndPost, directionalDiagnosis,
        evacuationAgreementCity, evacuationAgreementFacility, evacuationAgreementPhoneNumber, evacuationAgreementNameAndPost,
        callType, callCause,
        name, surname, patronymic, gender, birthDate, regPlace, jobPlace, insuranceCompany, insurancePolicy, passport,
        mainDiagnosis, complications, accompanying,
        driveResult, hospitalizationPlace,
        brigadeDoctor, additionalDoctor, paramedicInBrigade, nurseInBrigade, driverInBrigade, driveDistance, complains, anamnesis
    } = data;

    // Форматирование дат
    [callReceiveTime, callTransferTime, teamCallTime, arrivalTime, transportStartTime, transportEndTime, callEndTime] = [callReceiveTime, callTransferTime, teamCallTime, arrivalTime, transportStartTime, transportEndTime, callEndTime].map(date => {
        if (date === '') {
            return '';
        } else {
            return date.split(/\r?\n/)[0].split(/[-T]+/)[2] + '.' +
                date.split(/\r?\n/)[0].split(/[-T]+/)[1] + '.' +
                date.split(/\r?\n/)[0].split(/[-T]+/)[0] + ', ' +
                date.split(/\r?\n/)[0].split(/[-T]+/)[3];
        }
    });

    if (birthDate !== '') {
        birthDate = birthDate.split(/\r?\n/)[0].split(/[-]+/)[2] + '.' +
            birthDate.split(/\r?\n/)[0].split(/[-]+/)[1] + '.' +
            birthDate.split(/\r?\n/)[0].split(/[-]+/)[0];
    }

    // Форматирование чекбокса
    function isChecked(isChecked) {
        if (isChecked) {
            return { val: `☒`, opts: { align: 'center', textAlignment: 'top', sz: 24, fontFamily: "Times New Roman" } }
        } else {
            return { val: `☐`, opts: { align: 'center', textAlignment: 'top', sz: 24, fontFamily: "Times New Roman" } }
        }
    }

    // Создание Word файла
    let tempFilePath = tempfile('.docx');
    docx.setDocSubject('testDoc Subject');
    docx.setDocKeywords('keywords');
    docx.setDescription('test description');

    // Запись данных
    let table1 = [
        [{
            val: "I. Время (часы, минуты)",
            opts: {
                gridSpan: 8,
                align: 'center',
                tableColWidth: 4261,
                b: true,
                sz: 20,
                fontFamily: "Times New Roman"
            }
        }],
        [`Прием вызова`, `Передача вызова бригаде`, `Выезд бригады`, `Прибытие на место вызова`,
            `Начало транспортировки / убытие`, `Окончание транспортировки`, `Окончание вызова`, `Авиатранспорт`],
        [`${callReceiveTime}`, `${callTransferTime}`, `${teamCallTime}`, `${arrivalTime}`, `${transportStartTime}`,
            `${transportEndTime}`, `${callEndTime}`, isChecked(isAirTransport === 'on')
        ]
    ]
    let table1Style = {
        borders: true,
        tableColWidth: 4261,
        tableSize: 24,
        sz: 20,
        tableColor: "ada",
        tableAlign: "left",
        tableFontFamily: "Times New Roman"
    }
    let table2 = [
        [{
            val: "II. Откуда:",
            opts: {
                gridSpan: 4,
                align: 'center',
                b: true,
                sz: 20,
                fontFamily: "Times New Roman"
            }
        }],
        [`Населенный пункт`, `Учр-е здравоохранения, отделение`, `№ телефона`, `ФИО, должность`],
        [`${startCity}`, `${startHealthCareFacility}`, `${startPhoneNumber}`, `${startNameAndPost}`]
    ]
    let table2Style = {
        borders: true,
        tableColWidth: 8522,
        tableSize: 24,
        sz: 20,
        tableColor: "ada",
        tableAlign: "left",
        tableFontFamily: "Times New Roman"
    }
    let table3 = [
        [{
            val: "III. С кем согласовано место госпитализации при эвакуации:",
            opts: {
                gridSpan: 4,
                align: 'center',
                b: true,
                sz: 20,
                fontFamily: "Times New Roman"
            }
        }],
        [`Населенный пункт`, `Учр-е здравоохранения, отделение`, `№ телефона`, `ФИО, должность`],
        [`${evacuationAgreementCity}`, `${evacuationAgreementFacility}`, `${evacuationAgreementPhoneNumber}`, `${evacuationAgreementNameAndPost}`]
    ]
    let table4 = [
        [
            {
                val: `IV. Вызов:`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                    b: true
                }
            },
            {
                val: `первичный`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callType === 'primaryCall').val}`,
                opts: `${isChecked(callType === 'primaryCall').opt}`
            },
            {
                val: `повторный`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callType === 'repeatedCall').val}`,
                opts: `${isChecked(callType === 'repeatedCall').opt}`
            },
            {
                val: `попутный`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callType === 'passingCall').val}`,
                opts: `${isChecked(callType === 'passingCall').opt}`
            },
            {
                val: `вызов другой бригады`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callType === 'anotherBrigadeCall').val}`,
                opts: `${isChecked(callType === 'anotherBrigadeCall').opt}`
            },
            {
                val: `прочее`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callType === 'otherCall').val}`,
                opts: `${isChecked(callType === 'otherCall').opt}`
            },
        ]
    ]
    let table4Style = {
        borders: false,
        tableSize: 24,
        sz: 24,
        tableColor: "ada",
        tableAlign: "left",
        align: 'left',
        textAlignment: 'top',
        tableFontFamily: "Times New Roman"
    }
    let table5 = [
        [
            {
                val: `Повод к вызову:`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                    b: true
                }
            },
            {
                val: `консультация на месте`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callCause === 'onSiteConsultCause').val}`,
                opts: `${isChecked(callCause === 'onSiteConsultCause').opt}`
            },
            {
                val: `эвакуация`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callCause === 'evacuationCause').val}`,
                opts: `${isChecked(callCause === 'evacuationCause').opt}`
            },
            {
                val: `операция`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callCause === 'operationCause').val}`,
                opts: `${isChecked(callCause === 'operationCause').opt}`
            },
            {
                val: `прочее`,
                opts: {
                    sz: 24,
                    fontFamily: "Times New Roman",
                }
            },
            {
                val: `${isChecked(callCause === 'otherCause').val}`,
                opts: `${isChecked(callCause === 'otherCause').opt}`
            },
        ]
    ]

    let finalData = [
        {
            type: "text",
            val: 'Нижнетагильский филиал',
            lopt: { align: 'center' },
            opt: { font_size: 16, bold: true, font_face: 'Times New Roman' }
        },
        {
            type: "image",
            path: path.resolve(__dirname, 'static/img/printingImage.png'),
            lopt: { align: 'right' }
        },
        {
            type: "table",
            val: table1,
            opt: table1Style
        },
        { type: 'linebreak' },
        {
            type: "table",
            val: table2,
            opt: table2Style
        },
        { type: 'linebreak' },
        {
            type: "text",
            val: `НАПРАВИТЕЛЬНЫЙ ДИАГНОЗ: ${directionalDiagnosis}`,
            lopt: { align: 'left' },
            opt: { font_size: 11, bold: true, font_face: 'Times New Roman', underline: true }
        },
        { type: 'linebreak' },
        {
            type: "table",
            val: table3,
            opt: table2Style
        },
        { type: 'linebreak' },
        {
            type: "table",
            val: table4,
            opt: table4Style
        },
        { type: 'linebreak' },
        {
            type: "table",
            val: table5,
            opt: table4Style
        }
    ]
    docx.createByJson(finalData);

    // Завершение
    docx.on('finalize', function (written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
    });
    docx.on('error', function (err) {
        console.log(err);
    });

    // Запись
    res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
        'Content-disposition': 'attachment; filename=testdoc.docx'
    });
    docx.generate(res);
});

// Сервер на порту 8080
app.listen(8080);
// Сообщение о старте
console.log('Сервер стартовал!');