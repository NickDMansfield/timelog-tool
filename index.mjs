#! /usr/bin/env node
import xlsx from "async-xlsx";
import program from "commander";
import _ from "lodash";
import request from "request";
import fs from "fs";
// Steps
// Receive args (file loc, worksheet to use, line count, output loc)
// -- Note that file paths must use process.cwd as this is an npm module

function excelDate(excelDate) {
  return new Date((excelDate - (25567 + 1)) * 86400 * 1000);
}
program
  .version("0.0.1")
  .option("-J, --settingsJSON <settingsJSON>", "settings JSON file")
  .option("-S, --source <source>", "data source")
  .parse(process.argv);

// Load settings

let settings = {};
if (program.settingsJSON) {
  try {
    fs.readFile(
      process.cwd() + "/" + program.settingsJSON,
      "utf-8",
      function (err, settingsFileBuffer) {
        settings = JSON.parse(settingsFileBuffer);
        console.log(settings);
      }
    );
  } catch (exc) {
    console.log(exc);
  }
  // Load the file
  const filePath = process.cwd() + "/" + program.source;
  xlsx.parseFileAsync(filePath, {}, (workbookData) => {
    const wbData = JSON.parse(JSON.stringify(workbookData));
    // Result is a 2D array with an object with name/data props for each worksheet
    //console.log(JSON.stringify(wbData, 0, 2));
    //   console.log(workbookData);

    const sheetNames = _.map(wbData, "name");

    // Grab all the rows in each worksheet
    const allFetchedRows = {};
    _.map(sheetNames, (sheetname) => {
      const sheet = _.find(wbData, { name: sheetname });
      processSheet(sheet);
    });

    // because I am a lazy sob
    setInterval(function () {
      console.log(JSON.stringify(wbData));
      xlsx.buildAsync(wbData, null, (error, xlsxBuffer) => {
        console.log(xlsxBuffer);
        // TODO: Replace this with filepath when the formatting is fixed
        fs.writeFile(
          process.cwd() + "/zzz" + program.source,
          xlsxBuffer,
          () => {
            console.log("done");
            process.exit();
          }
        );
      });
    }, 10000);
    /*
     */

    // console.log(allFetchedRows);

    // Group comments and times if necessary

    // Exclusion logic - required/blacklisted

    // Append upload Id to each row being uploaded

    // Attempt to upload

    function processSheet(sheet, accountSettings) {
      const sheetData = sheet.data;
      const sheetDate = sheetData[0][6];
      let currentRowIndex = 1;
      let taskTitle;
      do {
        const row = sheetData[currentRowIndex];
        taskTitle = row[0];
        processRow(row, sheetDate, accountSettings);
        currentRowIndex++;
      } while (taskTitle);
    }

    function processRow(row, sheetDate = null, accountSettings) {
      const taskTitle = row[0];
      const existingRowTaskRecords = allFetchedRows[taskTitle];
      const startTime = excelDate(row[1]);
      const endTime = excelDate(row[2]);
      let taskDate = sheetDate ? excelDate(sheetDate) : null;
      // Because the plain date loads it up as the next day
      taskDate = new Date(taskDate.valueOf() - 30000);
      const parsedRow = {
        taskTitle,
        startTime,
        endTime,
        taskDescription: row[3],
        timeSplitMinutes: (endTime - startTime) / 60000,
        taskUploadId: row[9],
        taskDate,
      };
      if (
        parsedRow.endTime &&
        parsedRow.startTime &&
        parsedRow.taskTitle &&
        parsedRow.timeSplitMinutes &&
        !parsedRow.taskUploadId
      ) {
        postRow(parsedRow, row, accountSettings);
      }
      if (existingRowTaskRecords) {
        existingRowTaskRecords.push(parsedRow);
      } else {
        allFetchedRows[taskTitle] = [parsedRow];
      }
    }

    function postRow(X, rawRow, accountSettings) {
      const rowToPost = {
        hours: (X.timeSplitMinutes / 60).toString(),
        notes: X.taskTitle + ": " + X.taskDescription,
        project_id: accountSettings.projectId,
        spent_date: X.taskDate,
        task_id: accountSettings.taskId,
        user_id: accountSettings.userId,
      };
      // console.log(rowToPost);
      request.post(
        {
          url: "https://api.harvestapp.com/api/v2/time_entries",
          headers: {
            "User-Agent": "efficiency bot",
            "Harvest-Account-ID": `${accountSettings.accountId}`,
            Authorization: accountSettings.authorization,
          },
          json: rowToPost,
        },
        function (error, response, body) {
          if (!error) {
            X.taskUploadId = body.id;
            while (rawRow.length < 10) {
              rawRow.push("");
            }
            rawRow[9] = X.taskUploadId;
          } else {
            console.log(error);
          }
        }
      );
    }
  });
}
program.exit();
