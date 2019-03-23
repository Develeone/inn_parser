// SET THIS VARIABLES TO "TRUE" IF YOU WANT TO WATCH ALL LOGS
var debug = true;
var maximumDebug = true;



const https = require("https");
const XLSX = require('exceljs');
const readline = require("readline");
const fs = require('fs');

const ANTICAPTCHA_KEY = "a069673283b049fea2a4d1dfb281c192";
var solvedCaptchaTokens = [];

const THREADS_MAX_COUNT = 20;
var runningThreads = 0;

var personsCount = 0;
var checkedPersonsCount = 0;
var correctPersonsCount = 0;
var incorrectPersonsCount = 0;
var invalidPersonsCount = 0;

var checkedPersonsIndexes = [];
var checkingPersonsIndexes = [];

var currentFileName = "";
var finishedFiles = [];

var currentWorkbook = null;
var inputWorksheet = null;
var correctWorksheet = null;
var incorrectWorksheet = null;
var invalidWorksheet = null;

var lastNameRow = 3,          // Фамилия
    firstNameRow = 4,         // Имя
    middleNameRow = 5,        // Отчество
    birthdayRow = 12,         // Дата рождения
    passportSeriesRow = 7,    // Серия паспорта
    passportNumberRow = 8,    // Номер паспорта
    passportDateRow = 10,     // Дата выдачи паспорта
    resultRow = 11;           // Куда писать результат








selectFile();









function inputValues () {
    const r1 = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    r1.question('Колонка "Фамилия": ', function (test) {
        lastNameRow = test.toUpperCase();
        r1.question('Колонка "Имя": ', function (test) {
            firstNameRow = test.toUpperCase();
            r1.question('Колонка "Отчество": ', function (test) {
                middleNameRow = test.toUpperCase();
                r1.question('Колонка "Дата рождения": ', function (test) {
                    birthdayRow = test.toUpperCase();
                    r1.question('Колонка "Серия паспорта": ', function (test) {
                        passportSeriesRow = test.toUpperCase();
                        r1.question('Колонка "Паспорта": ', function (test) {
                            passportNumberRow = test.toUpperCase();
                            r1.question('Колонка "Дата выдачи паспорта": ', function (test) {
                                passportDateRow = test.toUpperCase();
                                r1.question('Колонка для записи результата: ', function (test) {
                                    resultRow = test.toUpperCase();
                                    r1.question('Количество потоков: ', function (test) {
                                        //THREADS_MAX_COUNT = test;

                                        startParsing();

                                        r1.close();
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    });
}


function selectFile() {

    if (maximumDebug)
        console.log("selectFile()");

    checkedPersonsCount = 0;
    correctPersonsCount = 0;
    incorrectPersonsCount = 0;
    invalidPersonsCount = 0;
    runningThreads = 0;
    checkedPersonsIndexes = [];
    checkingPersonsIndexes = [];
    currentWorkbook = null;
    inputWorksheet = null;
    correctWorksheet = null;
    incorrectWorksheet = null;
    invalidWorksheet = null;

    var inputFiles = fs.readdirSync("input");

    for (var fileIndex in inputFiles) {
        var _fileName = inputFiles[fileIndex];

        if (finishedFiles.indexOf(_fileName) < 0) {
            console.log("\n\nЗапуск обработки файла " + _fileName);

            startParsing(_fileName);
            finishedFiles.push(_fileName);
            currentFileName = _fileName;
            return;
        }
    }

    console.log("\n\n\nВсе файлы обработаны. Работа программы завершена");
}


function startParsing(fileName) {

    if (maximumDebug)
        console.log("Starting parsing");

    currentWorkbook = new XLSX.Workbook();

    currentWorkbook.xlsx.readFile('input/' + fileName)
        .then(function() {
            inputWorksheet = currentWorkbook.getWorksheet(1);

            personsCount = inputWorksheet.rowCount-1;

            correctWorksheet      = currentWorkbook.addWorksheet('Корректные',   {state: 'visible'});
            incorrectWorksheet    = currentWorkbook.addWorksheet('Некорректные', {state: 'visible'});
            invalidWorksheet      = currentWorkbook.addWorksheet('Неполные',     {state: 'visible'});

            copyRow(inputWorksheet, 1, correctWorksheet,   1);
            copyRow(inputWorksheet, 1, incorrectWorksheet, 1);
            copyRow(inputWorksheet, 1, invalidWorksheet,   1);

            runWebParser();
        });
}

function runWebParser() {
    console.log("\nTotal passports count: " + personsCount + "\n");

    process.stdout.cursorTo(0);
    process.stdout.write("Progress: 0/" + personsCount + ", correct: 0, incorrect: 0, invalid: 0");


    for (var i = 0; i < THREADS_MAX_COUNT; i++)
        runNewThread();

    var threadCreator = setInterval( function () {

        if (checkedPersonsCount == personsCount) {
            exportXlsx();
            clearInterval(threadCreator);
        }

        if (runningThreads < THREADS_MAX_COUNT && (personsCount - (checkedPersonsCount + runningThreads) > 0)) {
            runNewThread();
        }
    }, 1000);
}

function exportXlsx() {
    currentWorkbook.xlsx.writeFile("output/" + currentFileName)
        .then(function() {
            console.log("\n\nОбработка файла " + currentFileName + " завершена");

            selectFile();
        });
}

function runNewThread() {

    if (debug)
        console.log("\nNew thread! Summary: " + runningThreads + " threads");

    if (runningThreads > personsCount)
        return;

    runningThreads++;

    checkPerson();
}


// Working with CAPTCHA

function getCaptchaPage() {

    if (maximumDebug)
        console.log("GETTING CAPTCHA PAGE -----------------------------------------");

    var options = {
        host: 'service.nalog.ru',
        port: 443,
        path: '/static/captcha-dialog.html?aver=3.32.1&sver=4.32.17&pageStyle=GM2',
        method: 'GET',
        headers: {
            accept: 'application/json'
        }
    };

    https.get(options, function (res) {
        res.setEncoding("utf8");

        var content = "";

        res.on("data", function (chunk) {
            content += chunk;
        });

        res.on("end", function () {
            getCaptcha(content);
        });

    }).on('error', function (e) {
        process.stdout.cursorTo(0);
        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + " - Slow internet connection");

        if (debug)
            console.log("\nGot error: " + e.message);

        runningThreads--;
    });
}

function getCaptcha(html) {

    if (maximumDebug)
        console.log("EXTRACTING CAPTCHA FROM PAGE -------------------");

    var captchaToken = findCaptchaToken(html);

    if (!captchaToken)
        return;

    loadCaptchaBase64(captchaToken);
}

function findCaptchaToken(html) {
    var startingIndex = html.indexOf('name="captchaToken" value="');

    if (startingIndex < 0) {
        console.error("Cannot take CAPTCHA token!!!");

        //checkPerson("","");

        runningThreads--;

        return;

    }

    if (debug)
        console.error("FOUND CAPTCHA!!!");

    startingIndex += 27;

    var endingIndex = html.indexOf('"', startingIndex);

    return html.substring(startingIndex, endingIndex);
}

function loadCaptchaBase64(captchaToken) {
    https.get('https://service.nalog.ru/static/captcha.html?a='+captchaToken+'&version=1', function (response) {
        response.setEncoding('base64');
        var body = "";
        response.on('data', (data) => { body += data; });
        response.on('end', () => {

            if (maximumDebug)
                console.log("CAPTCHA BASE64: " + body);

            onCaptchaLoaded(body, captchaToken);
        });
    }).on('error', function (e) {
        process.stdout.cursorTo(0);
        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + " - Slow internet connection");

        if (debug)
            console.log("\nGot error 2: " + e.message);
        runningThreads--;
    });
}

function onCaptchaLoaded(base64, captchaToken) {

    if (maximumDebug)
        console.log("CREATING CAPTCHA SOLVERS TASK ----------------");

    var responseBody = "";

    var postData = JSON.stringify({
        "clientKey": ANTICAPTCHA_KEY,
        "task": {
            "type": "ImageToTextTask",
            "body": base64,
            "phrase": "false",
            "case": "false",
            "numeric": 1,
            "math": 0,
            "minLength": 6,
            "maxLength": 6
        }
    });

    var options = {
        hostname: 'api.anti-captcha.com',
        port: 443,
        path: '/createTask',
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        }
    };

    var req = https.request(options, (res) => {
        res.on('data', (data) => { responseBody += data});
        res.on('end', () => {

            if (debug)
                console.log("CAPTCHA task creation response: " + responseBody);

            onTaskIdGot(JSON.parse(responseBody).taskId, captchaToken);
        });
    });

    req.on('error', (e) => {
        process.stdout.cursorTo(0);
        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + " - Slow internet connection");

        if (debug)
            console.log("\nGot error 3: " + e.message);
        runningThreads--;
    });

    req.write(postData);
    req.end();
}

function onTaskIdGot(taskId, captchaToken) {
    if (!taskId) {
        console.error("Error while creating ANTICAPTCHA task!");
        runningThreads--;
    }

    if (debug)
        console.log("Task ID got: "+taskId);

    setTimeout(() => {
        var interval = setInterval ( () => {

            var responseBody = "";

            var postData = JSON.stringify({
                "clientKey": ANTICAPTCHA_KEY,
                "taskId": taskId
            });

            var options = {
                hostname: 'api.anti-captcha.com',
                port: 443,
                path: '/getTaskResult',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            };

            var req = https.request(options, (res) => {
                res.on('data', (data) => { responseBody += data});
                res.on('end', () => {
                    // console.log(responseBody);

                    var response = JSON.parse(responseBody);

                    if (response.status == "ready") {
                        clearInterval(interval);

                        if (solvedCaptchaTokens.indexOf(captchaToken) >= 0)
                            return;

                        var captcha = response.solution.text;

                        solvedCaptchaTokens.push(captchaToken);

                        gotCaptchaCode(captcha, captchaToken);
                    }
                    else if (response.status != "processing") {
                        console.error("\nError while solving captcha! \n" + responseBody);
                    }
                });
            });

            req.on('error', (e) => {
                process.stdout.cursorTo(0);
                process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + " - Slow internet connection");

                if (debug)
                    console.error("\nError 3: ", e);
            });

            req.write(postData);
            req.end();

        }, 1000);
    }, 5000);
}

function gotCaptchaCode(captcha, captchaToken) {
    if (debug)
        console.log("\nGot new CAPTCHA: " + captcha + " with captchaToken: " + captchaToken);

    postCaptcha(captcha, captchaToken);
}

function postCaptcha(captcha, captchaToken) {
    var postData = {
        'captcha': captcha,
        'captchaToken': captchaToken
    };

    var options = {
        hostname: 'service.nalog.ru',
        port: 443,
        path: '/static/captcha-proc.json',
        method: 'POST',
        headers: {
            'accept': '*/*',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'en-US,en;q=0.8',
            'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    };

    var responseBody = "";

    var req = https.request(options, (res) => {
        if (debug)
            console.log("Posting CAPTCHA key to special page");

        res.on('data', (data) => { responseBody += data });
        res.on('end', () => {
            responseBody = JSON.parse(responseBody);

            if (debug) {
                console.log("Got captcha key " + responseBody);
            }

            if (responseBody.ERRORS) {
                runningThreads--;
                return;
            }

            checkPerson("", responseBody);
        });
        res.on('error', () => {
            console.log("GOT ERROR GETTING CAPTCHA KEY!!!");
            runningThreads--;
        });
    });

    req.on('error', (e) => {
        runningThreads--;

        console.error("\nError while posting captcha key to the special form!!! " + e);
    });

    req.write(serialize(postData));
    req.end();
}

// Working with persons

function checkPerson(captcha, captchaToken) {
    var responseBody = "";

    var targetRowIndex = 0;

    for (var i = 2; i <= personsCount+1; i++) {
        var existsInChecked = false;
        var existsInChecking = false;

        for (var id in checkedPersonsIndexes) {
            if (i == checkedPersonsIndexes[id]) {
                existsInChecked = true;
                break;
            }
        }

        for (var id in checkingPersonsIndexes) {
            if (i == checkingPersonsIndexes[id]) {
                existsInChecking = true;
                break;
            }
        }

        if (!existsInChecked && !existsInChecking) {
            targetRowIndex = i;
            break;
        }
    }

    if (!targetRowIndex) {
        if (debug)
            console.error("\nPerson undefined");
        runningThreads--;
        return;
    }

    checkingPersonsIndexes.push(targetRowIndex);

    if (!getCell(targetRowIndex, passportSeriesRow)) {
        checkedPersonsCount++;
        invalidPersonsCount++;

        removeArrayItem(checkingPersonsIndexes, targetRowIndex);
        checkedPersonsIndexes.push(targetRowIndex);

        copyRow(inputWorksheet, targetRowIndex, invalidWorksheet, invalidPersonsCount+1);

        invalidWorksheet.getRow(invalidPersonsCount+1).getCell(resultRow).value = "passport_series: \"Отсутствует серия паспорта\"";

        runningThreads--;

        process.stdout.cursorTo(0);
        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount);

        return;
    } else {
        if ((getCell(targetRowIndex, passportSeriesRow) + "").length != 4) {
            checkedPersonsCount++;
            invalidPersonsCount++;

            removeArrayItem(checkingPersonsIndexes, targetRowIndex);
            checkedPersonsIndexes.push(targetRowIndex);

            copyRow(inputWorksheet, targetRowIndex, invalidWorksheet, invalidPersonsCount+1);

            invalidWorksheet.getRow(invalidPersonsCount+1).getCell(resultRow).value = "passport_series: \"Некорректная серия паспорта\"";

            runningThreads--;

            process.stdout.cursorTo(0);
            process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount);

            return;
        }
    }

    var postData = {
        'c': 'innMy',
        'fam': getCell(targetRowIndex, lastNameRow),
        'nam': getCell(targetRowIndex, firstNameRow),
        'otch': getCell(targetRowIndex, middleNameRow),
        'bdate': getDate(getCell(targetRowIndex, birthdayRow)),
        'bplace': "",
        'doctype': '21',
        'docno': (getCell(targetRowIndex, passportSeriesRow)+"").slice(0, 2) + " " + (getCell(targetRowIndex, passportSeriesRow)+"").slice(2) + " " + getCell(targetRowIndex, passportNumberRow),
        'docdt': "",
        'captcha': captcha,
        'captchaToken': captchaToken
    };

    if (debug)
        console.log(serialize(postData));


    if (maximumDebug)
        console.log(postData);

    function getDate(date){

        if (!date)
            return date;

        if (date.value)
            if (typeof date.value.getMonth === 'function')
                if (date.value.getFullYear() > 1900 && date.value.getFullYear() < 2019) {
                    var day = date.value.getDate();
                    var month = date.value.getMonth()+1;
                    return (day < 10 ? '0' : '') + day + '.'
                        + (month < 10 ? '0' : '') + month + '.'
                        + date.value.getFullYear();
                }

        try {
            var _date = date.split("/");
        } catch (e) {
            return date;
        }

        if (_date.length == 1)
            _date = date.split(".");

        var year = _date[2];
        var month = "0" + _date[1];
        var day = "0" + _date[0];

        if (year.length == 2) {
            if (year[0] == "0")
                year = "20" + year;
            else
                year = "19" + year;
        }

        if (month > 12) {
            var _day = day;
            day = month;
            month = _day;
        }

        _date = day.substr(-2) + '.' + month.substr(-2) + '.' + year;

        return _date;
    }

    var options = {
        hostname: 'service.nalog.ru',
        port: 443,
        path: '/inn-proc.do',
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    };

    var req = https.request(options, (res) => {
        res.on('data', (data) => { responseBody += data; if (maximumDebug) console.log(data); });
        res.on('end', () => {

            if (debug)
                console.log("Ответ формы nalog.ru: " + responseBody);

            responseBody = JSON.parse(responseBody);

            if (responseBody.code == 0 || responseBody.code == 1) {
                checkedPersonsCount++;

                if (responseBody.code) {
                    correctPersonsCount++;
                    copyRow(inputWorksheet, targetRowIndex, correctWorksheet, correctPersonsCount+1);
                    correctWorksheet.getRow(correctPersonsCount+1).getCell(resultRow).value = responseBody.inn;
                } else {
                    incorrectPersonsCount++;
                    copyRow(inputWorksheet, targetRowIndex, incorrectWorksheet, incorrectPersonsCount+1);
                }

                checkedPersonsIndexes.push(targetRowIndex);

                if (debug) {
                    console.log("\nPerson checked, passport " + (responseBody.code ? "" : "in") + "correct." + (responseBody.code ? " INN: " + responseBody.inn : ""));
                }
            }
            else {
                if (responseBody.ERRORS) {
                    if (!responseBody.ERRORS.captcha) {
                        checkedPersonsCount++;
                        invalidPersonsCount++;

                        removeArrayItem(checkingPersonsIndexes, targetRowIndex);
                        checkedPersonsIndexes.push(targetRowIndex);

                        copyRow(inputWorksheet, targetRowIndex, invalidWorksheet, invalidPersonsCount+1);

                        invalidWorksheet.getRow(invalidPersonsCount+1).getCell(resultRow).value = responseBody.ERRORS;

                        if (checkedPersonsCount > 3 && checkedPersonsCount == invalidPersonsCount)
                            console.log("\nЛибо в файле очень много некорректных данных, либо столбцы указаны неверно. Пожалуйста, проверьте, всё ли в порядке.");
                    } else {
                        removeArrayItem(checkingPersonsIndexes, targetRowIndex);

                        process.stdout.cursorTo(0);
                        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + "         - Вводится CAPTCHA");

                        getCaptchaPage();

                        return;
                    }
                } else {

                    if (responseBody.code) {
                        if (responseBody.code == 99) {

                            removeArrayItem(checkingPersonsIndexes, targetRowIndex);
                            runningThreads--;

                            return;
                        }
                    }

                    checkedPersonsCount++;
                    invalidPersonsCount++;

                    checkedPersonsIndexes.push(targetRowIndex);

                    // TODO: person.error = {error: "Неизвестная ошибка"};

                    console.log("\nНеизвестная ошибка! Пользователь добавлен в список \"неполные данные\". \nТекст ошибки: ", responseBody);
                }
            }

            process.stdout.cursorTo(0);
            process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount);

            removeArrayItem(checkingPersonsIndexes, targetRowIndex);
            runningThreads--;
        });
    });

    req.on('error', (e) => {
        removeArrayItem(checkingPersonsIndexes, targetRowIndex);
        runningThreads--;

        process.stdout.cursorTo(0);
        process.stdout.write("Progress: " + checkedPersonsCount + "/" + personsCount + ", correct: " + correctPersonsCount + ", incorrect: " + incorrectPersonsCount + ", invalid: " + invalidPersonsCount + "         - Slow internet connection");

        if (debug)
            console.error("\nError!!! " + e);
    });

    if (debug)
        console.log("Отправляем данные в форму nalog.ru: "/*, postData*/);

    req.write(serialize(postData));
    req.end();
}


function serialize (obj) {
    var str = [];
    for (var p in obj)
        if (obj.hasOwnProperty(p)) {
            str.push(encodeURIComponent(p) + "=" + encodeURIComponent(obj[p]));
        }
    return str.join("&");
}

function removeArrayItem(arr) {
    var what, a = arguments, L = a.length, ax;
    while (L > 1 && arr.length) {
        what = a[--L];
        while ((ax= arr.indexOf(what)) !== -1) {
            arr.splice(ax, 1);
        }
    }
    return arr;
}

function getCell(rowIndex, cellIndex) {
    if (!currentWorkbook)
        return " ";


    if (maximumDebug)
        console.log(currentWorkbook.getWorksheet(1).getRow(rowIndex).getCell(cellIndex));

    return currentWorkbook.getWorksheet(1).getRow(rowIndex).getCell(cellIndex);
}

function copyRow(fromSheet, fromRowIndex, toSheet, toRowIndex) {
    toSheet.getRow(toRowIndex).style = fromSheet.getRow(fromRowIndex).style;
    toSheet.getRow(toRowIndex).height = fromSheet.getRow(fromRowIndex).height;

    for (var i = 1; i <= fromSheet.columnCount; i++) {
        toSheet.getRow(toRowIndex).getCell(i).value = fromSheet.getRow(fromRowIndex).getCell(i).value;
        toSheet.getRow(toRowIndex).getCell(i).style = fromSheet.getRow(fromRowIndex).getCell(i).style;

        toSheet.getColumn(i).style  = fromSheet.getColumn(i).style;
        toSheet.getColumn(i).width  = fromSheet.getColumn(i).width;
    }
}