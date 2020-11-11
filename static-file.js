// Express, tempfile, officegen
let express = require('express');
let tempfile = require('tempfile');
let officegen = require('officegen');
let docx = officegen('docx');
let path = require('path');

// Express-приложение
let app = express();
// Подключение статичных файлов
app.use("/static", express.static('./static/'));
// Маршрут для главной страницы
app.get('/', function (req, res) {
    res.sendFile('index.html', {root: __dirname});
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
    const {
        mainDate, cardNumber, paramedicOnCall,
        callReceiveTime, callTransferTime, teamCallTime, arrivalTime, transportStartTime, transportEndTime, callEndTime, isAirTransport,
        startCity, startHealthCareFacility, startPhoneNumber, startNameAndPost, directionalDiagnosis,
        evacuationAgreementCity, evacuationAgreementFacility, evacuationAgreementPhoneNumber, evacuationAgreementNameAndPost,
        call, callCause,
        name, surname, patronymic, gender, birthDate, regPlace, jobPlace, insuranceCompany, insurancePolicy, passport,
        mainDiagnosis, complications, accompanying,
        driveResult, hospitalizationPlace,
        brigadeDoctor, additionalDoctor, paramedicInBrigade, nurseInBrigade, driverInBrigade, driveDistance, complains, anamnesis
    } = data;

    // Создание Word файла
    let tempFilePath = tempfile('.docx');
    docx.setDocSubject('testDoc Subject');
    docx.setDocKeywords('keywords');
    docx.setDescription('test description');

    // Запись данных
    let pObj = docx.createP({align: 'center'});
    pObj.addText('Нижнетагильский филиал', {font_size: 16, bold: true, font_face: 'Times New Roman'});
    pObj.addImage(path.resolve(__dirname, 'static/img/printingImage.png'), {align: 'right', x:0});
    // pObj.addText(`${mainDate}, ${cardNumber}`, {bold: true, underline: true});

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