const { GoogleSpreadsheet } = require('google-spreadsheet');
const fs = require('fs');
const http = require("http");
const functions = require('@google-cloud/functions-framework');
const nodemailer = require('nodemailer');
const {google} = require('googleapis');



// Initialize the sheet - doc ID is the long id in the sheets URL
const doc = new GoogleSpreadsheet('1YLVHUpvnXZ5KwvqTXTqD8vs7uTkAaDqVs9JRDe1JyVk');


const CREDENTIALS = JSON.parse(fs.readFileSync('./clear-gantry-388204-b5be11dc86aa.json'));

const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];


const getAnniversaryData = async () => {
   // Initialize Auth - see https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
    
        client_email: CREDENTIALS.client_email,
        private_key: CREDENTIALS.private_key
    });

    await doc.loadInfo(); // loads document properties and worksheets

    const sheet1 = doc.sheetsByIndex[0];
    let rows = await sheet1.getRows();

    let currentDate = new Date();
    let currentMonth = currentDate.getMonth();
    let currentDay = currentDate.getDate();
    let currentDayWithMonth = `${currentDay}/${currentMonth}`;
    let data = [];
    let flag = false;
    for(let i = 0; i < rows.length; i++) {

        let date = new Date(rows[i]['Marriage Date']);
        let dateWithMonth = `${date.getDate()}/${date.getMonth()}`;
        
        if(currentDayWithMonth === dateWithMonth) {

            const diffInMilliseconds = Math.abs(currentDate - date);
            const millisecondsInYear = 1000 * 60 * 60 * 24 * 365.25; 
            const anniversaryInYears = Math.floor(diffInMilliseconds / millisecondsInYear);

            data.push(new Person(rows[i].Name, rows[i].DOB, rows[i].Email, rows[i]['Marriage Date'], anniversaryInYears));
            flag = true;
        }
        
    }

    if(flag) {
        console.log(data);
        return data;
    } else {
        console.log("Hello");
        return "No data found to display";
    }


}




function Person(Name, DOB, Email, Marr, Anv) {
    this.name = Name;
    this.dob = DOB;
    this.email = Email;
    this.marriageDate = Marr;
    this.anniversary = Anv;
}



functions.http('helloHttp', async (req, res) => {

    let data = await getAnniversaryData();
    
    if(req.url == "/anni") {

        for(let i = 0; i < data.length; i++) {
            sendMailToClient(data[i])
            .then(result => console.log("Mail is sent", result))
            .catch(err => console.log(err.message));
            
        }

        res.send();
        
    } else {
        res.send("<h1>Page Not Found!</h1>")
    }
});


 

const CLIENT_ID = '1068914366323-9vaoadmqgggki0g0t7k5400biad087hb.apps.googleusercontent.com'
const CLIENT_SECRET = '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'
const REDIRECT_URI = 'https://developers.google.com/oauthplayground'
const REFRESH_TOKEN = '1//046I3DFuvzY0-CgYIARAAGAQSNwF-L9IrprLJ65IcP5AegJVkQkvhUcaXaAJFOyyPa8vefwGw_n1TUGRCCjvIuW7uP7al_Y1yy-o'

const oAuth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);
oAuth2Client.setCredentials({refresh_token: REFRESH_TOKEN});

async function sendMailToClient(data) {

    try {

        const accessToken = await oAuth2Client.getAccessToken();
        const transport = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                type: 'OAuth2',
                user: 'indraindrani1999@gmail.com',
                clientId: CLIENT_ID,
                clientSecret: CLIENT_SECRET,
                refreshToken: REFRESH_TOKEN,
                accessToken: accessToken
            }
        });

        const mailOptions = {
            from: 'Indrajit <indraindrani1999@gmail.com>',
            // to: data.email,
            to: "educatorindrajit@gmail.com",
            subject: `Happy ${data.anniversary}th Anniversary from ABC Company!`,
            text: `Dear ${data.name},

            Happy ${data.anniversary}th anniversary! We're delighted to celebrate this special milestone with you as a valued customer of ABC Company. Your support means the world to us, and we're honored to be a part of your journey.
                    
            Thank you for choosing our "Holiday Home" services. We hope that our accommodations have provided you with unforgettable experiences and moments of joy. Your continued trust inspires us to keep delivering exceptional service. Here's to ${data.anniversary} years of cherished memories and to many more incredible anniversaries ahead! We're grateful for the opportunity to serve you and look forward to creating more wonderful moments together.
                  
         Congratulations again, and thank you for being a part of the ABC Company family.
                              
         Best wishes,
         Indrajit Paul
         ABC Company`
          ,
            
        }

        const result = await transport.sendMail(mailOptions);
        return result;

    } catch(error) {
        return error;
    }

    
}


