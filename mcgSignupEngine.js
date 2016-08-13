////////////////////////////////////////////////////////////////////////////////////////
// This script is intended to take the results from a Google form for clinic signups and 
// randomize people to clinic dates in an attempt to create more fair signups than the 
// previous first-come-first-served system. 
//
// The algorithm gathers all the people who signed up and their available dates
// and then goes through the dates in ascending order of the number of volunteers
// available that date and assigns people to that date. It then prints out the results 
// in a new sheet titled "Clinic Schedule".
//
// This script could be more object-oriented to clean things up a bit. I don't have time
// now, but might go through some later time to change that. Who knows. There is also no
// testing.
//
// It's also dependent on a specific format. Dates must be in a mm/dd/yy format.
// Column B must be a list of names, Column C can be a list of associated emails or it 
// could be some other associated text, and Column D must contain comma-separated lists
// of the dates that people are available. This must all be on a sheet named
// "Form Responses 1".
//
// Created by Hutton Brandon (hbrandon@augusta.edu). Don't hesitate to contact if you
// have questions.
////////////////////////////////////////////////////////////////////////////////////////

//
// Variables
//

var ss = SpreadsheetApp.getActiveSpreadsheet();
var formResponses = ss.getSheetByName('Form Responses 1');
var lastRow = formResponses.getLastRow();

// Number of people per clinic dates derived from box on sheet
var numVolunteers = Math.floor(formResponses.getRange('A2').getValues()[0][0]);
  
// Create array for volunteers from the entries
var peopleArray = new Array();

// Create array for all clinic dates
var clinicDates = new Array();

// Create a nx3 array of all responses
var entriesArray = formResponses.getRange('B2:D' + lastRow).getValues();

//
// Create menu for custom functions
//
function onOpen() {
  var menu = [{name: 'Create suggested schedule sheet', functionName: 'deleteScheduleAndCreateNewOne'}];
  ss.addMenu('MCG Signup Engine', menu);
};

//
// Main Functions
//

function createScheduleSheet () {
  
  // Randomize people to dates as volunteers and alternates
  assignVolunteersToDate();
  
  // Create new sheet listing the dates with volunteers and waitlist
  var scheduleSheet = ss.insertSheet('Clinic Schedule');
  ss.setActiveSheet(scheduleSheet);
  addStaticItems();
  addClinicDatesObjects();
  addPeopleArrayWaitlist();
};

function deleteClinicSheet() {
  var sht = ss.getSheetByName('Clinic Schedule');
  
  // If a 'Clinic Schedule' sheet exists, delete it. Else log lack of existence.
  (sht) ? ss.deleteSheet(sht) : Logger.log("No Clinic Schedule tab exists");
};

function deleteScheduleAndCreateNewOne() {
  deleteClinicSheet();
  createScheduleSheet();
};

//
// Supporting Functions
//

// Turn each entry into a Javascript object
function populatePeopleArray() {
  for (i=0; i<entriesArray.length; i++) {
    var datesArray = entriesArray[i][2].split(', ');
    for (j=0; j<datesArray.length; j++) { datesArray[j] = toDate(datesArray[j]) };
    peopleArray.push( {
      name: entriesArray[i][0],
      email: entriesArray[i][1],
      dates: datesArray
    });
  }; 
  Logger.log('populatePeopleArray successful');
};

// Populate clinicDates in order
function populateClinicDates() {
  for (i=0; i<entriesArray.length; i++) {
    datesString = entriesArray[i][2];
    arr = datesString.split(", ");
    for (j=0; j<arr.length; j++) {
      var target = toDate(arr[j]);
      var found = false;
      
      // Crappy 'includes' function
      for (k=0; k<clinicDates.length; k++) {
        if (clinicDates[k].date.getTime() === target.getTime()) {
          found = true;
          clinicDates[k].signups += 1;
          break;
        };
      };
      // If not included in clinicDates, add it to that array
      if (found == false) {
        clinicDates.push({
          date: target,
          volunteers: [],
          signups: 1 // To keep track of how many signed up for that date for sorting purposes
        });
      };   
    };
  };
  
  // Sort by the number of signups, lowest to highest
  var clinicDatesBySignups = clinicDates.slice(0);
  clinicDatesBySignups.sort(function(a,b) {return a.signups - b.signups});
  clinicDates = clinicDatesBySignups;
  
  //for (i=0; i<clinicDates.length; i++) {Logger.log(clinicDates[i].date)};
  Logger.log('populateClinicDates successful');
  Logger.log(clinicDates);
};

// Convert mm/dd/yy to datetime
function toDate(dateStr) {
    var parts = dateStr.split("/");
    return new Date("20" + parts[2], parts[0] - 1, parts[1]);
};

// Check if dates are the same
function areSameDate(a,b) { 
  if (a.getTime() === b.getTime()) {
    return true;
  } else { 
    return false
  };
};

// Convert datetime to mm/dd
function toFormattedDate(date) {
  var day = new Date(date);
  Logger.log(day);
  return (day.getMonth()+1) + '/' + day.getDate();
};

//
// Function to randomize people to dates as volunteers
function assignVolunteersToDate() {
// Turn each entry into a Javascript object
  populatePeopleArray();
  
  // Populate clinicDates in order
  populateClinicDates();
  
  //
  // Assign people to dates
  
  
  // For each date in clinicDates, find all the people that have listed that date as available and add each name to an array.
  for (i=0; i<clinicDates.length; i++) {
    var possibleNames = [];
    var currentObject = clinicDates[i];
    var currentDate = currentObject.date;
    
    for (j=0; j<peopleArray.length; j++) {
      var name = peopleArray[j].name;
      var email = peopleArray[j].email;
      var availableDates = peopleArray[j].dates;
      
      for (k=0; k<availableDates.length; k++) {
        //Logger.log(currentDate.getTime() + ' vs ' + availableDates[k].getTime());
        if (areSameDate(currentDate, availableDates[k])) {
          //Logger.log(currentDate + ' = ' + availableDates[k]);
          possibleNames.push({
            name: name,
            email: email,
            score: Math.random()
          });
        };
      };
      //Logger.log(currentDate + " ++> ");
      
    };
    
    // Sort by highest random score
    possibleNames = possibleNames.sort(function(a,b) {return a.score - b.score});
    
    // Pare it down to the top #numVolunteers
    var length = possibleNames.length;
    while (length > numVolunteers) {
      possibleNames.splice(0,1);
      length--;
    };
    
    for (m=0; m<possibleNames.length; m++) {
      var currentPerson = possibleNames[m];
      // Add those people to currentObject in clinicDates
      currentObject.volunteers.push(currentPerson);
      
      // And remove them from the peopleArray using email addresses
      peopleArray = peopleArray.filter( function(object) { return object.email !== currentPerson.email });
    };
  };
};

// Read contents of clinicDates array
function readClinicDates() {
  for (i=0; i<clinicDates.length; i++) {
    //Logger.log(clinicDates[i]);
  };
};



// Add static text and items to the active sheet
function addStaticItems() {
  var active = ss.getActiveSheet();
  
  // Add name of clinic
  active.getRange('A1').setValue('[Clinic Name] Summer 2016 Schedule')
        .setFontSize(16)
        .setFontWeight('bold');
  // Color name of clinic background
  active.getRange('A1:E1').setBackgroundRGB(33, 157, 220);
  
  // Add labels
  active.getRange('A5').setValue('Dates')
                       .setBackgroundRGB(255, 130, 0)
                       .setFontWeight('bold')
                       .setFontSize(12);
  active.getRange('A12').setValue('Waitlist')
                        .setFontWeight('bold')
                        .setBackgroundRGB(0, 220, 0);
  active.getRange('B12').setValue('Dates Available')
                        .setFontWeight('bold')
                        .setBackgroundRGB(0, 220, 0);
  
  // Set C3:Z3 to have center alignment for date values
  active.getRange('C5:Z5').setHorizontalAlignment('center');
};

//
// Add clinicDates objects to active sheet
function addClinicDatesObjects() {
  var active = ss.getActiveSheet();
  
  // Sort clinicDates by date
  var clinicDatesByDate = clinicDates.splice(0);
  clinicDatesByDate.sort(function(a,b) {return a.date - b.date});
  Logger.log('Sorted by dates');
  Logger.log(clinicDatesByDate);
  
  
  for (i=0;i<clinicDatesByDate.length;i++) {
    var currentObject = clinicDatesByDate[i];
    var volunteers = currentObject.volunteers;
    var length = volunteers.length;
    // Selects a 1xlenght+1 range
    var range = active.getRange(5,i+2,length+1);
    var date = currentObject.date;
    var values = [[date]];
    
    // Resize column
    var column = range.getColumn();
    active.setColumnWidth(column, 120);
    
    // Set colors
    var dateCell = active.getRange(5,i+2).setBackgroundRGB(255, 130, 0)
                                         .setFontWeight('bold');
    
    // Set color of volunteer cells if non-zero number of volunteers
    if (length > 0) {
      var volunteerCells = active.getRange(6,i+2,length);
      volunteerCells.setBackgroundRGB(204, 204, 255);
      volunteerCells.setWrap(true);
    };
    
    // Set wraps
    var wraps = []
    for (n=0; n<length+1; n++) {
      wraps.push([true]);
    };
    range.setWraps(wraps);
    
    for (j=0; j<volunteers.length; j++) {
      values.push([volunteers[j].name + ' ' + volunteers[j].email]);
    };
    range.setValues(values);              
  };
};

// Add remainder of peopleArray as a waitlist to active sheet
function addPeopleArrayWaitlist() {
  var active = ss.getActiveSheet();
  var height = peopleArray.length;
  
  for (i=0; i<height; i++) {
    var currentObject = peopleArray[i];
    var name = currentObject.name;
    var email = currentObject.email;
    
    // Fumbling around to get dates in the right format and in a string
    var dates = currentObject.dates;
    dates = dates.toString();
    Logger.log(dates);
    var dateStrArr = dates.split(',');
    for (j=0;j<dateStrArr.length; j++) {
      dateStrArr[j] = toFormattedDate(dateStrArr[j]);
    };
    dates = dateStrArr.join(', ');
    
    
    Logger.log(dates);
    // Select a 1x2 section on bottom left side of sheet
    var cell1 = active.getRange(13 + i, 1);
    cell1.setValue(name + ' ' + email);
    cell1.setWrap(true)
         .setBackgroundRGB(0, 220, 0);
    var cell2 = active.getRange(13 + i, 2);
    cell2.setValue(dates);
    cell2.setWrap(true)
         .setBackgroundRGB(0, 220, 0)
         .setNumberFormat('@STRING@');
  };
};

//
// Testing
//

//
// Main Test Function

function runTestSuite() {
  
  numVolunteersIsANumber();
  
};

// Tests
function numVolunteersIsANumber() {
  if ((typeof(numVolunteers) === 'number') && (numVolunteers > 0)) {
    Logger.log("PASS: Number of volunteers per date is specified correctly as a natural number");
  } else {
    Logger.log("FAIL: Number of volunteers per date is not a natural number");
  };
};
    
    
  
