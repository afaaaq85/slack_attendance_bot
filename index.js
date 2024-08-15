// const { App, ExpressReceiver } = require("@slack/bolt");
// const ExcelJS = require("exceljs");
// require("dotenv").config();
// const fs = require("fs");
// const { google } = require("googleapis");
// const apikeys = require("./apikey.json");
// const moment = require("moment-timezone");

// //authorize with google
// const SCOPE = ["https://www.googleapis.com/auth/drive"];

// const authorizeGoogleConnection = async () => {
//   const jwtClient = new google.auth.JWT(apikeys.client_email, null, apikeys.private_key, SCOPE);
//   await jwtClient.authorize();
//   return jwtClient;
// };
// const uploadFile = async (authClient) => {
//   return new Promise(async (resolve, reject) => {
//     const drive = google.drive({ version: "v3", auth: authClient });
//     const fileName = "employees_attendance.xlsx";
//     const folderId = process.env.FOLDER_ID; // Folder ID where the file is located

//     let fileId = null;

//     const fileMetaData = {
//       name: fileName,
//       parents: [folderId],
//     };

//     try {
//       // Search for the file by name in the specified folder
//       const searchResponse = await drive.files.list({
//         q: `name='${fileName}' and '${folderId}' in parents`,
//         fields: "files(id, name)",
//         spaces: "drive",
//       });

//       // If the file exists, get its ID
//       if (searchResponse.data.files.length > 0) {
//         fileId = searchResponse.data.files[0].id;
//       }
//     } catch (error) {
//       return reject("Error searching for file: " + error.message);
//     }

//     const media = {
//       body: fs.createReadStream(fileName),
//       mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//     };

//     try {
//       if (fileId) {
//         // Update the existing file
//         const updateResponse = await drive.files.update({
//           fileId: fileId,
//           media: media,
//         });
//         resolve(updateResponse.data);
//       } else {
//         // Create a new file if it doesn't exist
//         const createResponse = await drive.files.create({
//           resource: fileMetaData,
//           media: media,
//           fields: "id",
//         });
//         resolve(createResponse.data);
//       }
//     } catch (error) {
//       reject("Error uploading file: " + error.message);
//     }
//   });
// };

// const receiver = new ExpressReceiver({
//   signingSecret: process.env.SIGNING_SECRET,
// });

// const app = new App({
//   token: process.env.SLACK_TOKEN,
//   receiver,
// });

// const employees = {
//   Afaq: {},
//   Ammar: {},
//   Nouman: {},
//   Aalee: {},
//   Hadi: {},
// };

// app.message(async ({ message, context }) => {
//   const userId = message.user;
//   const text = message.text;
//   const timestamp = parseFloat(message.ts);
//   const messageTime = moment(timestamp * 1000)
//     .tz("Asia/Karachi")
//     .format("HH:mm:ss");

//   const currentDate = moment(timestamp * 1000)
//     .tz("Asia/Karachi")
//     .format("YYYY-MM-DD");

//   try {
//     const userInfo = await app.client.users.info({
//       token: context.botToken,
//       user: userId,
//     });

//     const userName = userInfo.user.name;

//     console.log("employee:", userName);
//     console.log("msg:", text);
//     console.log("time:", messageTime);

//     if (userName === "afaq.codrivity" && ["in", "reached", "online"].includes(text.toLowerCase())) {
//       if (!employees["Afaq"][currentDate]) {
//         employees["Afaq"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Afaq"][currentDate]["In Time"] = messageTime;
//       await updateExcelFile();
//     }
//     if (
//       userName === "afaq.codrivity" &&
//       ["out", "leaving", "left", "offline"].includes(text.toLowerCase())
//     ) {
//       if (!employees["Afaq"][currentDate]) {
//         employees["Afaq"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Afaq"][currentDate]["Out Time"] = messageTime;
//       await updateExcelFile();
//     }

//     //for Ammar
//     if (userName === "afaqatofficial" && ["in", "reached", "online"].includes(text.toLowerCase())) {
//       if (!employees["Ammar"][currentDate]) {
//         employees["Ammar"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Ammar"][currentDate]["In Time"] = messageTime;
//       await updateExcelFile();
//     }
//     if (
//       userName === "afaqatofficial" &&
//       ["out", "leaving", "left", "offline"].includes(text.toLowerCase())
//     ) {
//       if (!employees["Ammar"][currentDate]) {
//         employees["Ammar"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Ammar"][currentDate]["Out Time"] = messageTime;
//       await updateExcelFile();
//     }

//     //for nouman
//     if (userName === "mughalfasih75" && ["in", "reached", "online"].includes(text.toLowerCase())) {
//       if (!employees["Nouman"][currentDate]) {
//         employees["Nouman"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Nouman"][currentDate]["In Time"] = messageTime;
//       await updateExcelFile();
//     }
//     if (
//       userName === "mughalfasih75" &&
//       ["out", "leaving", "left", "offline"].includes(text.toLowerCase())
//     ) {
//       if (!employees["Nouman"][currentDate]) {
//         employees["Nouman"][currentDate] = { "In Time": "", "Out Time": "" };
//       }
//       employees["Nouman"][currentDate]["Out Time"] = messageTime;
//       await updateExcelFile();
//     }
//   } catch (error) {
//     console.error("Error handling message:", error);
//   }
// });

// async function updateExcelFile() {
//   const filePath = "employees_attendance.xlsx";
//   const workbook = new ExcelJS.Workbook();

//   if (fs.existsSync(filePath)) {
//     await workbook.xlsx.readFile(filePath);
//   } else {
//     workbook.addWorksheet("Attendance");
//   }

//   const worksheet = workbook.getWorksheet("Attendance");

//   // Ensure the columns are correctly set
//   worksheet.columns = [
//     { header: "Employee", key: "employee", width: 20 },
//     { header: "Date", key: "date", width: 15 },
//     { header: "In Time", key: "inTime", width: 30 },
//     { header: "Out Time", key: "outTime", width: 30 },
//     { header: "Break Start", key: "breakStart", width: 30 },
//     { header: "Break End", key: "breakEnd", width: 30 }, 
//   ];

//   for (const [employee, dates] of Object.entries(employees)) {
//     for (const [date, times] of Object.entries(dates)) {
//       let rowFound = false;

//       // Search for the existing row with the same employee and date
//       worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
//         if (row.getCell("employee").value === employee && row.getCell("date").value === date) {
//           row.getCell("inTime").value = times["In Time"];
//           row.getCell("outTime").value = times["Out Time"];
//           rowFound = true;
//         }
//       });

//       if (!rowFound) {
//         console.log("Row not found for employee:", employee, "and date:", date);
//         worksheet.addRow({
//           employee,
//           date,
//           inTime: times["In Time"],
//           outTime: times["Out Time"],
//         });
//       }
//     }
//   }

//   await workbook.xlsx.writeFile(filePath);
//   // authorizeGoogleConnection().then(uploadFile).catch(console.error);
// }

// receiver.router.get("/status", (req, res) => {
//   res.status(200).send("Server is running!");
// });

// (async () => {
//   await app.start(process.env.PORT);
//   console.log("⚡️ Slack Bolt app is running!");
// })();


const { App, ExpressReceiver } = require("@slack/bolt");
const ExcelJS = require("exceljs");
require("dotenv").config();
const fs = require("fs");
const { google } = require("googleapis");
const apikeys = require("./apikey.json");
const moment = require("moment-timezone");

// Authorize with Google
const SCOPE = ["https://www.googleapis.com/auth/drive"];
const authorizeGoogleConnection = async () => {
  const jwtClient = new google.auth.JWT(apikeys.client_email, null, apikeys.private_key, SCOPE);
  await jwtClient.authorize();
  return jwtClient;
};

const uploadFile = async (authClient) => {
  return new Promise(async (resolve, reject) => {
    const drive = google.drive({ version: "v3", auth: authClient });
    const fileName = "employees_attendance.xlsx";
    const folderId = process.env.FOLDER_ID;

    let fileId = null;

    const fileMetaData = {
      name: fileName,
      parents: [folderId],
    };

    try {
      const searchResponse = await drive.files.list({
        q: `name='${fileName}' and '${folderId}' in parents`,
        fields: "files(id, name)",
        spaces: "drive",
      });

      if (searchResponse.data.files.length > 0) {
        fileId = searchResponse.data.files[0].id;
      }
    } catch (error) {
      return reject("Error searching for file: " + error.message);
    }

    const media = {
      body: fs.createReadStream(fileName),
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    };

    try {
      if (fileId) {
        const updateResponse = await drive.files.update({
          fileId: fileId,
          media: media,
        });
        resolve(updateResponse.data);
      } else {
        const createResponse = await drive.files.create({
          resource: fileMetaData,
          media: media,
          fields: "id",
        });
        resolve(createResponse.data);
      }
    } catch (error) {
      reject("Error uploading file: " + error.message);
    }
  });
};

const receiver = new ExpressReceiver({
  signingSecret: process.env.SIGNING_SECRET,
});

const app = new App({
  token: process.env.SLACK_TOKEN,
  receiver,
});

const employees = {
  Afaq: {},
  Ammar: {},
  Nouman: {},
  Aalee: {},
  Hadi: {},
};

// Mapping between Slack usernames and employee names in the employees object
const userMapping = {
  "afaq.codrivity": "Afaq",
  "afaqatofficial": "Ammar",
  "mughalfasih75": "Nouman",
  "aalee_username": "Aalee",  // Update with the actual Slack username
  "hadi_username": "Hadi",    // Update with the actual Slack username
};

app.message(async ({ message, context }) => {
  const userId = message.user;
  const text = message.text.toLowerCase();
  const timestamp = parseFloat(message.ts);
  const messageTime = moment(timestamp * 1000)
    .tz("Asia/Karachi")
    .format("hh:mm:ss A");

  const currentDate = moment(timestamp * 1000)
    .tz("Asia/Karachi")
    .format("YYYY-MM-DD");

  try {
    const userInfo = await app.client.users.info({
      token: context.botToken,
      user: userId,
    });

    const userName = userInfo.user.name;
    const employeeName = userMapping[userName];  // Map the username to the employee name

    if (employeeName) {
      if (!employees[employeeName][currentDate]) {
        employees[employeeName][currentDate] = {
          "In Time": "",
          "Out Time": "",
          "Break Start": "",
          "Break End": "",
        };
      }

      if (["in", "reached", "online"].includes(text)) {
        employees[employeeName][currentDate]["In Time"] = messageTime;
      }

      if (["out", "leaving", "left", "offline"].includes(text)) {
        employees[employeeName][currentDate]["Out Time"] = messageTime;
      }

      if (["break start", "taking break"].includes(text)) {
        employees[employeeName][currentDate]["Break Start"] = messageTime;
      }

      if (["break end", "back from break"].includes(text)) {
        employees[employeeName][currentDate]["Break End"] = messageTime;
      }

      await updateExcelFile();
    } else {
      console.warn(`No mapping found for user: ${userName}`);
    }
  } catch (error) {
    console.error("Error handling message:", error);
  }
});

async function updateExcelFile() {
  const filePath = "employees_attendance.xlsx";
  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(filePath)) {
    await workbook.xlsx.readFile(filePath);
  } else {
    workbook.addWorksheet("Attendance");
  }

  const worksheet = workbook.getWorksheet("Attendance");

  worksheet.columns = [
    { header: "Employee", key: "employee", width: 20 },
    { header: "Date", key: "date", width: 15 },
    { header: "In Time", key: "inTime", width: 30 },
    { header: "Out Time", key: "outTime", width: 30 },
    { header: "Break Start", key: "breakStart", width: 30 },
    { header: "Break End", key: "breakEnd", width: 30 },
  ];

  for (const [employee, dates] of Object.entries(employees)) {
    for (const [date, times] of Object.entries(dates)) {
      let rowFound = false;

      worksheet.eachRow({ includeEmpty: true }, (row) => {
        if (row.getCell("employee").value === employee && row.getCell("date").value === date) {
          row.getCell("inTime").value = times["In Time"];
          row.getCell("outTime").value = times["Out Time"];
          row.getCell("breakStart").value = times["Break Start"];
          row.getCell("breakEnd").value = times["Break End"];
          rowFound = true;
        }
      });

      if (!rowFound) {
        worksheet.addRow({
          employee,
          date,
          inTime: times["In Time"],
          outTime: times["Out Time"],
          breakStart: times["Break Start"],
          breakEnd: times["Break End"],
        });
      }
    }
  }

  await workbook.xlsx.writeFile(filePath);
  authorizeGoogleConnection().then(uploadFile).catch(console.error);
}

receiver.router.get("/status", (req, res) => {
  res.status(200).send("Server is running!");
});

(async () => {
  await app.start(process.env.PORT);
  console.log("⚡️ Slack Bolt app is running!");
})();
