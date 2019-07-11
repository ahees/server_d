const fs = require("fs");
const start= new Date().getTime();
const generateDocx = require('generate-docx');


const arr = {
"DogovorDate": "09-07-2018",
"DogovorNomer": "5е46456",
"LikarName": "КУЛАЙ Наталія Іванівна",
"LikarNameShort": "КУЛАЙ Н. І.",
"ManagerName": "ЧЕРВІНСЬКА Маргарита Андріївна",
"ManagerNameShort": "ЧЕРВІНСЬКА М. А.",
"PaketCode": 1199,
"PaketName": "Оплодотворение in vitro (ОИВ) с использованием донорских ооцитов (Донор)",
"​PatientAddress": "Україна, Київ, Тульчинська 9-Б, 68, під'їзд 2, поверх 4",
"PatientBirthDate": "01-01-1990",
"PatientEmail": "roman.kamenskyy@eleks.com",
"PatientId": 1,
"PatientName": "Eleks Irina Petrovna",
"PatientName1": "Eleks",
"PatientName2": "Irina",
"PatientName3": "Petrovna",
"PatientPassport": "NJ 12345",
"PatientPhone": "380(66)705-82-19;380(44)297-12-51;380(50)312-54-78"
};

function validate(val) {
    if ((val === undefined) || (val === null)) {
        return ' '
    }
    else {
        return val
    }
}


function fio (p,n,b) {
    n = n.charAt(0); 
    if (b.length>1) {
        b = b.charAt(0) + '.';
    }
    return `${p} ${n}. ${b}`
}

function strToArr (str) {
    var arr = str.split(';');
    return arr
}

// console.log(data);
var objdata = arr;
const options = {
    template: {
        filePath: 'file/input/dogovor.docx',
        data: {}
    },
    save: {
        filePath: 'file/output/dogovor.docx'
    }
};

var elem = options.template.data;
    for(key in objdata){

        elem[key] = validate(objdata[key]);
}
elem.PatientNameShort = fio(elem.PatientName1, elem.PatientName2, elem.PatientName3);
let phone = strToArr(elem.PatientPhone)
elem.PatientPhone1 = phone[0];

elem.PatientBirthDate = elem.PatientBirthDate.replace(/\-/g,'.');

//console.log(elem);


generateDocx(options, (error, message) => {
    if (error) {
      console.error(error)
    } else {
      console.log(message);
      require('child_process').execSync(`docto -f "file/output/dogovor.docx" -O "file/output/dogovor.pdf" -T wdFormatPDF  -OX .pdf`, (err, stdout, stderr) => {
        
        if (err) {
          console.error(err);
          return;
        }

        

      });
      const end = new Date().getTime();
        console.log(`SecondWay: ${end - start}ms`);
      
    }
  });

  