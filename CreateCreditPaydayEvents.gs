/**
 * Google Apps Script for Automating Google Calendar Events from Google Sheets
 *
 * Description: This script automates the creation of Google Calendar events
 * from data within a Google Sheet.  It processes rows, creates
 * events, and writes the event status and ID back to the sheet.
 *
 * Features:
 * -   Reads event data from a sheet named 'Calendar Events Data'.
 * -   Handles event titles, descriptions, start/end times, and attendees.
 * -   Skips rows where events have already been created (based on Event ID).
 * -   Writes event creation status and Event IDs back to the sheet.
 * -   Adds a custom menu to the Google Sheet for easy execution.
 *
 * Usage:
 * 1.  Ensure your Google Sheet is named 'Calendar Events Data' and has the
 * following headers in the first row: 'Event Title', 'Event Description',
 * 'Start Time', 'End Time', 'Attendees' (optional), 'Status', 'Event ID'.
 * 2.  Copy and paste this script into the Google Apps Script editor.
 * 3.  Run the 'createEventsFromSheet' function from the custom menu
 * 'Calendar Actions' in your Google Sheet.
 *
 * Author: Artemio B. Gallardo
 * YouTube Channel:
 *
 * License: Apache License, Version 2.0
 * (http://www.apache.org/licenses/LICENSE-2.0)
 *
 * Disclaimer:
 * This script is provided as-is, without any warranty.  Use at your own risk.
 */


const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Acciones de Calendario')
    .addItem('Crear Eventos Para Pagar Creditos', 'createEventsFromSheet')
    .addToUi();
};

const createEventsFromSheet = () => {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Dias de corte creditos');
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet named 'Calendar Events Data' not found.");
    return;
  }

  const dataRange = sheet.getDataRange(); // Process all rows in the sheet
  const [head, ...data] = dataRange.getValues();

  const getColumnIndex = (columnName) => {
    const index = head.indexOf(columnName);
    if (index === -1) {
      SpreadsheetApp.getUi().alert(`Column '${columnName}' not found in the header row.`);
      throw new Error(`Column '${columnName}' not found.`);
    }
    return index;
  };

  try {
    const titleCol = getColumnIndex('Titulo');
    const descriptionCol = getColumnIndex('Descripción');
    const cutDayCol = getColumnIndex('Día de corte');
    const daysToPayCol = getColumnIndex('Dias para pago');
    const attendeesCol = head.indexOf('Invitados'); // Optional
    const statusCol = getColumnIndex('Estatus');
    const idCol = getColumnIndex('ID Evento');
    const fechaPagoCol = getColumnIndex('Fecha Pago');

    const calendar = CalendarApp.getDefaultCalendar();
    const output = [];

    data.forEach((row, i) => {
      const title = row[titleCol];
      const description = row[descriptionCol];
      const cutDay = row[cutDayCol];
      const daysToPay = row[daysToPayCol];
      const attendees = attendeesCol !== -1 ? row[attendeesCol] : '';
      const existingEventId = row[idCol]; // Read the event ID from the sheet for this row

      Logger.log(`Row ${i + 2}: cutDay=${cutDay}, daysToPay=${daysToPay}`);

      let statusToSet = "";
      // Initialize idToSet with the existing ID.
      // This ensures if skipped, the existing ID is maintained.
      let idToSet = existingEventId;
  
      const result = validateProperties(i, existingEventId, title, cutDay, daysToPay);
      Logger.log(result.statusToSet, result.idToSet);
      
      let dateToPayStart = 0
      if (result.isValid) {
        const eventOptions = {
          description: description || '',
          sendInvites: false
        };

        if (attendees) {
          eventOptions.guests = attendees;
        }

        dateToPayStart = convertNumberToDate(cutDay, daysToPay)

        try {
          const event = calendar.createAllDayEvent(title, dateToPayStart, eventOptions);
          calendar.
          idToSet = event.getId(); // Successfully created, so use new ID
          statusToSet = "Created (All Day)";
          Logger.log('Created event with id:', idToSet);
        } catch (error) {
          statusToSet = `Error: ${error.message}`; // Use error.message for a cleaner status
          idToSet = ""; // Error during creation, so no ID
          Logger.log(`Error creating event in row ${i + 2}:`, error);
        }
      }

      output.push([statusToSet, idToSet, dateToPayStart]);
    });

    // Write the status and event ID back to the sheet
    if (output.length > 0) { // Only write if there's output data
      const startRow = 2; // Start from the second row (after the header)
      // Prepare ranges for batch setting values for efficiency
      const statusValues = output.map(o => [o[0]]);
      const idValues = output.map(o => [o[1]]);
      const payDate = output.map(o => [o[2]]);

      sheet.getRange(startRow, statusCol + 1, output.length, 1).setValues(statusValues);
      sheet.getRange(startRow, idCol + 1, output.length, 1).setValues(idValues);
      sheet.getRange(startRow, fechaPagoCol + 1, output.length, 1).setValues(payDate);
    }

    SpreadsheetApp.getUi().alert('Calendar events creation process completed.');

  } catch (error) {
    Logger.log("Error during script execution:", error);
    SpreadsheetApp.getUi().alert(`An error occurred: ${error.message}`);
  }
};

const convertNumberToDate = (day, offset = 0, monthOffset = 0) => {
  // Get today's date to know the current month and year
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + monthOffset; // allow shifting to next/previous month

  // Create a new Date object using the given day
  let resultDate = new Date(year, month, day);

  Logger.log(`Cut Date: ${resultDate}`);

  // Validate: if the day is out of range (e.g., 31 in February), JS will roll it over
  if (resultDate.getDate() !== day) {
    throw new Error(`Invalid day ${day} for month ${month + 1}`);
  }

  // Apply offset (e.g., +10 days)
  resultDate.setDate(resultDate.getDate() + offset);

  // Set to one second before midnight 
  resultDate.setHours(00, 00, 00, 01);
   
  Logger.log(`Pay Date: ${resultDate}`);

  return resultDate;
};

const validateProperties = (i, existingEventId, title, cutDay, daysToPay) => {
  let statusToSet = "";
  let idToSet = "";
  let isValid = true;

  Logger.log(`Inside validateProperties`);
  if (existingEventId) {
    statusToSet = "Skipped: Event already created.";
    idToSet = existingEventId; 
    isValid = false;
    Logger.log(`Skipping row ${i + 2}: Event already created (ID: ${existingEventId})`);
  } else if (!title || !cutDay || !daysToPay) {
    statusToSet = "Error: Missing title, start time, or end time.";
    idToSet = ""; 
    isValid = false;
    Logger.log(`Error creating event in row ${i + 2}: Missing required fields [title, cutDay, daysToPay].`);
  } 

  return { statusToSet, idToSet, isValid };
};
