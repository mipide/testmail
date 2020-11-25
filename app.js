var nodemailer = require('nodemailer');
const Excel = require('exceljs');



var transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: 'mp.detolle@gmail.com',
      pass: 'fannyaud18'
    }
  });
  
  var mailOptions = {
    from: 'mp.detolle@gmail.com',
    to: 'mp.detolle@airede.com',
    subject: 'Sending Email using Node.js',
    text: 'That was easy!'
  };
  
  var filename ="test.xlsx"
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile(filename).then(function(){
        var workSheet = workbook.worksheets[0]
        workSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            currRow = workSheet.getRow(rowNumber); 
            mailOptions.to=currRow.getCell(2).text
            

            currRow.getCell(3).value=new Date();
            currRow.commit();
            console.log("User Name :" + currRow.getCell(2).text +", senddate :" +currRow.getCell(3).value);
           

            
            transporter.sendMail(mailOptions, function(error, info){
                if (error) {
                  console.log(error);
                } else {
                  console.log('Email sent: ' + info.response);
                  console.log("previous date =" + currRow.getCell(3).value);
                  currRow.getCell(3).value=new Date();
                  currRow.commit();
                  console.log("User Name :" + currRow.getCell(2).text +", senddate :" +currRow.getCell(3).value);
                }
              });
              workbook.xlsx.writeFile(filename);  
        });
        
    });

  
  
/*
  workbook.xlsx.writeFile(filename);


  
  */