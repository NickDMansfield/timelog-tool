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
        if (err) {
          console.log(err);
        }
        settings = JSON.parse(settingsFileBuffer);
        console.log(settings);
        // Load the file
        const filePath = process.cwd() + "/" + program.source;
        xlsx.parseFileAsync(filePath, {}, (workbookData) => {
          const wbData = JSON.parse(JSON.stringify(workbookData));
          // Result is a 2D array with an object with name/data props for each worksheet
          //console.log(JSON.stringify(wbData, 0, 2));
          //   console.log(workbookData);
          _.map(settings.accountsToUse, (mappedAccountSetting) => {
            const dateNotes = {};
            const dateStrings = [];
            const workRowsByDate = {};
            const sheetNames = _.map(wbData, "name");

            // Grab all the rows in each worksheet
            const allFetchedRows = {};
            _.map(sheetNames, (sheetname) => {
              const sheet = _.find(wbData, { name: sheetname });
              processSheet(sheet, mappedAccountSetting);
            });

            // Since lumped accounts didn't post earlier, we handle them now
            if (shouldGroup(mappedAccountSetting)) {
              _.map(dateStrings, (dateString) => {
                const rowsToProcess = workRowsByDate[dateString];
                const totalDayMinutes = _.sum(
                  _.map(rowsToProcess, (rtp) => rtp.timeSplitMinutes)
                );
                const roundingHoursForGrouping =
                  getRoundingHoursForGrouping(mappedAccountSetting);
                // Combine the entire day into a single row
                const rowToPost = makePostableRow(
                  mappedAccountSetting.lumpSettings.title || "General Work",
                  // We don't always want to round grouped hours
                  roundingHoursForGrouping
                    ? roundWorkMinutes(
                        totalDayMinutes,
                        roundingHoursForGrouping * 60
                      )
                    : totalDayMinutes,
                  "\r\n" + dateNotes[dateString],
                  rowsToProcess[0].taskDate,
                  mappedAccountSetting
                );
                //  console.log(rowToPost);
                postRow(rowToPost, null, mappedAccountSetting);
              });
            }

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
              const taskDescription = row[3];
              let taskDate = sheetDate ? excelDate(sheetDate) : null;
              // Because the plain date loads it up as the next day
              taskDate = new Date(taskDate.valueOf() - 30000);
              const parsedRow = {
                taskTitle,
                startTime,
                endTime,
                taskDescription,
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
                if (!isRowExcluded(parsedRow, accountSettings)) {
                  // If there is a lump setting, we don't want itemized rows. Instead we want a single lumped count
                  const dateString = taskDate.toDateString();
                  if (dateStrings.indexOf(dateString) < 0) {
                    dateStrings.push(dateString);
                  }
                  if (
                    dateNotes[dateString] === null ||
                    dateNotes[dateString] === undefined
                  ) {
                    dateNotes[dateString] = "";
                  }
                  dateNotes[dateString] +=
                    taskTitle + ":" + taskDescription + "\r\n";
                  if (shouldGroup(accountSettings)) {
                    // IDK if I like this inside or outside the if. It will need to be moved if we want this info later
                    if (workRowsByDate[dateString]) {
                      workRowsByDate[dateString].push(parsedRow);
                    } else {
                      workRowsByDate[dateString] = [parsedRow];
                    }
                  } else {
                    postRow(parsedRow, row, accountSettings);
                  }
                }
              }
              if (existingRowTaskRecords) {
                existingRowTaskRecords.push(parsedRow);
              } else {
                allFetchedRows[taskTitle] = [parsedRow];
              }
            }

            function postRow(X, rawRow, accountSettings) {
              request.post(
                {
                  url: "https://api.harvestapp.com/api/v2/time_entries",
                  headers: {
                    "User-Agent": "efficiency bot",
                    "Harvest-Account-ID": `${accountSettings.accountId}`,
                    Authorization: accountSettings.authorization,
                  },
                  json: X,
                },
                function (error, response, body) {
                  //  console.log(response);
                  if (!error) {
                    X.taskUploadId = body.id;
                    if (rawRow != null) {
                      while (rawRow.length < 10) {
                        rawRow.push("");
                      }
                      rawRow[9] = X.taskUploadId;
                    }
                  } else {
                    console.log(error);
                  }
                }
              );
            }
          });
        });
      }
    );
  } catch (exc) {
    console.log(exc);
  }
}

function roundWorkMinutes(_minutes, _lumpSizeInMinutes) {
  return Math.round(_minutes / _lumpSizeInMinutes) * _lumpSizeInMinutes;
}

function makePostableRow(
  taskTitle,
  timeSplitMinutes,
  taskDescription,
  taskDate,
  accountSettings
) {
  const rowToPost = {
    hours: (timeSplitMinutes / 60).toString(),
    notes: taskTitle + ": " + taskDescription,
    project_id: accountSettings.projectId,
    spent_date: taskDate,
    task_id: accountSettings.taskId,
    user_id: accountSettings.userId,
  };
  return rowToPost;
}

function shouldGroup(accountSettings) {
  return (
    accountSettings.lumpSettings &&
    ((accountSettings.lumpSettings.hours &&
      accountSettings.lumpSettings.hours > 0) ||
      (accountSettings.lumpSettings.title &&
        accountSettings.lumpSettings.title.length > 0))
  );
}

function getRoundingHoursForGrouping(accountSettings) {
  return accountSettings?.lumpSettings?.hours > 0
    ? accountSettings?.lumpSettings?.hours
    : null;
}

function isRowExcluded(parsedRow, accountSettings) {
  // Check for excluded ticket names
  if (accountSettings.excludedTicketNames) {
    for (let ticketNameToExclude of accountSettings.excludedTicketNames) {
      if (ticketNameToExclude.length >= 0) {
        let stringToSearch = ticketNameToExclude;
        let stringAsteriskIndex = stringToSearch.indexOf("*");
        let endIndex =
          stringToSearch.length > 0 ? stringToSearch.length - 1 : 0;
        if (stringAsteriskIndex >= 0) {
          // here we check for wildcard matches
          endIndex = stringAsteriskIndex;
        }
        stringToSearch = stringToSearch.substring(0, endIndex);
        console.log("string:" + stringToSearch);
        console.log(parsedRow.taskTitle);
        if (
          stringToSearch.length > 0 &&
          parsedRow.taskTitle
            .toLowerCase()
            .indexOf(stringToSearch.toLowerCase()) >= 0
        ) {
          return true;
        }
      }
    }
  }
  return false;
}
