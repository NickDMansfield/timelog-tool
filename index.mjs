#! /usr/bin/env node
import xlsx from "async-xlsx";
import xlsx2 from "xlsx";
import program from "commander";
import _ from "lodash";
import uuid from "node-uuid";
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

// Load the file
xlsx.parseFileAsync(
  process.cwd() + "/" + program.source,
  {},
  (workbookData) => {
    const wbData = JSON.parse(JSON.stringify(workbookData));
    // Result is a 2D array with an object with name/data props for each worksheet
    //console.log(JSON.stringify(wbData, 0, 2));
    //   console.log(workbookData);

    const sheetNames = _.map(wbData, "name");

    // Grab all the rows in each worksheet
    const allFetchedRows = {};
    _.map(sheetNames, (sheetname) => {
      const sheet = _.find(wbData, { name: sheetname });
      const sheetData = sheet.data;
      let currentRowIndex = 1;
      let taskTitle;
      do {
        taskTitle = sheetData[currentRowIndex][0];
        console.log(taskTitle);
        const existingRowTaskRecords = allFetchedRows[taskTitle];
        const startTime = excelDate(sheetData[currentRowIndex][1]);
        const endTime = excelDate(sheetData[currentRowIndex][2]);
        const parsedRow = {
          taskTitle,
          startTime,
          endTime,
          taskDescription: sheetData[currentRowIndex][3],
          timeSplitMinutes: (endTime - startTime) / 60000,
          taskUploadId: sheetData[currentRowIndex][9],
        };
        if (
          parsedRow.endTime &&
          parsedRow.startTime &&
          parsedRow.taskTitle &&
          parsedRow.timeSplitMinutes
        ) {
          postRow(parsedRow);
        }
        if (existingRowTaskRecords) {
          existingRowTaskRecords.push(parsedRow);
        } else {
          allFetchedRows[taskTitle] = [parsedRow];
        }
        currentRowIndex++;
      } while (taskTitle);
    });
    // console.log(allFetchedRows);

    // Group comments and times if necessary

    // Exclusion logic - required/blacklisted

    // Append upload Id to each row being uploaded

    // Attempt to upload

    const origRows = allFetchedRows["ORIG-709"];
    //  console.log(origRows);
    //console.log(origRows.length);

    function processRow(row) {
      taskTitle = row[0];
      console.log(taskTitle);
      const existingRowTaskRecords = allFetchedRows[taskTitle];
      const startTime = excelDate(row[1]);
      const endTime = excelDate(row[2]);
      const parsedRow = {
        taskTitle,
        startTime,
        endTime,
        taskDescription: row[3],
        timeSplitMinutes: (endTime - startTime) / 60000,
        taskUploadId: row[9],
      };
      postRow(parsedRow);
      if (parsedRow.notes !== "null: null") {
        if (existingRowTaskRecords) {
          existingRowTaskRecords.push(parsedRow);
        } else {
          allFetchedRows[taskTitle] = [parsedRow];
        }
      }
    }
    function postRow(X, cb) {
      const rowToPost = {
        hours: (X.timeSplitMinutes / 60).toString(),
        notes: X.taskTitle + ": " + X.taskDescription,
        project_id: 30547106,
        spent_date: "2021-11-12",
        task_id: 15357383,
        user_id: 3992099,
      };
      console.log(rowToPost);
      request.post(
        {
          url: "https://api.harvestapp.com/api/v2/time_entries",
          headers: {
            "User-Agent": "efficiency bot",
            "Harvest-Account-ID": "1356984",
            Authorization:
              "Bearer 2839898.pt.NGdsmSna_8KIliNdDO8AlkVTmjUKrDy8VnH9sEz6X17jYfYdyxbCg1l23IJY6r4nyooRMtBEwPwZzKS2piVouQ",
          },
          json: rowToPost,
        },
        function (error, response, body) {
          if (!error) {
            X.taskUploadId = body.id;
          }
          console.log(X);
        }
      );
    }
  }
);
// If it succeeds, save the upload Id onto the file
