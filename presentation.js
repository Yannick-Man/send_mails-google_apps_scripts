/*
 * A Google app script that send your email data from a google spreadsheet.
 *
 * Copyright (c) 2011,2012 Emergya
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * Author: Francisco Perez <fperez@emergya.com>
 *
 */

// Presentation functions

function send() {
  
    // default spredsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

   // sheet emails
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet1'));
  var sheet = spreadsheet.getActiveSheet();
       
  
  var startRow = 1;  // First row of data to process 
  var numRows = 1; // Number of rows to send

  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, 1, numRows, 4)

  // Recorremos la hoja de calculo PESTANIACAL y enviamos los emails correspondientes.
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
  var row = data[i];
 
  var emailAddress=row[2]; // cc : list of emails
  var subject=row[0]; // subject email
  var description=row[1]; // descripcion email

      MailApp.sendEmail(emailAddress, subject, description);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
       
  }

  var app = UiApp.getActiveApplication();
  //app.close();

// The following line is REQUIRED for the widget to actually close.
  return app;
}


