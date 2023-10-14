const express = require('express');
const router = express.Router();
const xlsx = require("xlsx")

const scheduleSE = "files/se.xlsx"
const scheduleEconomic = "files/Economic.xlsx"
let year = 3
var isEconomicFaculty = false
var faculty, major, importRange, economicMajors

function takeEconomicMajor(text) {
    let matches = text.match(/«([^»]+)»/g);
    return matches.map(match => match.slice(1, -1));
}

function readExcelFile(file) {
    let wb = xlsx.readFile(file)
    let sheetValue = wb.Sheets[wb.SheetNames[0]]
    faculty = sheetValue['A6'].v.split(' ').slice(1).join(' ')
    if (isEconomicFaculty) {
        economicMajors = takeEconomicMajor(sheetValue['A7'].v)
    } else {
        major = sheetValue['A7'].v.split(' ').slice(1).join(' ').match(/"([^"]+)"/)[1]
    }
    return xlsx.utils.sheet_to_json(sheetValue, {range: importRange, header: 10})
}

function separateMajors(obj, result) {
    const text = obj["Дисципліна, викладач"]

    var isEconomic = /\(екон\.|ек\.|ек|економ\. теор\.\)/.test(text)
    var isFinances = /\(фін\.|фінанси\)/.test(text)
    var isMarketing = /\(мар\.|маркетинг\.|мар\)/.test(text)
    var isManagement = /\(мен\.|менеджмент\.|мен\)/.test(text)

    if (isEconomic) {
        obj["Спеціальність"] = "Економіка"
        result[0].push(obj)
    }
    if (isFinances) {
        obj["Спеціальність"] = "Фінанси, банківська справа та страхування"
        result[1].push(obj)
    }
    if (isMarketing) {
        obj["Спеціальність"] = "Маркетинг"
        result[2].push(obj)
    }
    if (isManagement) {
        obj["Спеціальність"] = "Менеджмент"
        result[3].push(obj)
    }
}

function parseSchedule(excelData) {
    var result = []
    var resultEconomic = [[], [], [], []]
    var day, time
    excelData.forEach(obj => {
        if (obj.hasOwnProperty("День")) {
            day = obj["День"]
        } else {
            obj["День"] = day
        }
        if (obj.hasOwnProperty("Час")) {
            time = obj["Час"]
        } else {
            obj["Час"] = time
        }
        if (obj.hasOwnProperty("Дисципліна, викладач")) {
            var newObj = {
                "Факультет": faculty,
                "Спеціальність": major,
                "Курс": year,
                "Дисципліна, викладач": obj["Дисципліна, викладач"],
                "Група": obj["Група"],
                "Тижні": obj["Тижні"],
                "День": obj["День"],
                "Час": obj["Час"],
                "Аудиторія": obj["Аудиторія"]
            };
            if (isEconomicFaculty) {
                var aud = obj["Ауд."]
                if (aud === "Д") {
                    aud = "Дистанційно"
                }
                newObj["Аудиторія"] = aud
                separateMajors(newObj, resultEconomic)
            } else {
                result.push(newObj)
            }
        }
    });
    if (isEconomicFaculty) {
        return resultEconomic
    } else {
        return result
    }
}

//розклад іпз
router.get('/se', function (req, res, next) {
    isEconomicFaculty = false
    importRange = "A10:F73";
    let excelDataSE = readExcelFile(scheduleSE)
    res.json(parseSchedule(excelDataSE))
});

//розклад ФАКУЛЬТКТУ економіки
router.get('/faculty/economics', function (req, res, next) {
    isEconomicFaculty = true
    importRange = "A10:F119";
    let excelDataEconomic = readExcelFile(scheduleEconomic)
    res.json(parseSchedule(excelDataEconomic))
});

//розклад економіки
router.get('/economic', function (req, res, next) {
    isEconomicFaculty = true
    importRange = "A10:F119";
    let excelDataEconomic = readExcelFile(scheduleEconomic)
    res.json(parseSchedule(excelDataEconomic)[0])
});

//розклад фінансів
router.get('/finances', function (req, res, next) {
    isEconomicFaculty = true
    importRange = "A10:F119";
    let excelDataEconomic = readExcelFile(scheduleEconomic)
    res.json(parseSchedule(excelDataEconomic)[1])
});

//розклад маркетингу
router.get('/marketing', function (req, res, next) {
    isEconomicFaculty = true
    importRange = "A10:F119";
    let excelDataEconomic = readExcelFile(scheduleEconomic)
    res.json(parseSchedule(excelDataEconomic)[2])
});

//розклад менеджменту
router.get('/management', function (req, res, next) {
    isEconomicFaculty = true
    importRange = "A10:F119";
    let excelDataEconomic = readExcelFile(scheduleEconomic)
    res.json(parseSchedule(excelDataEconomic)[3])
});

module.exports = router;
