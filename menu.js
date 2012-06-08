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

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [];
  menuEntries.push({name: "Send email generic", functionName: "sendEmail"});
  ss.addMenu("Send", menuEntries);
}


function sendEmail () {
  send();
}
