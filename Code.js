/***************************************************
 * נתוני חיבור ל-Fireberry (התאם לערכים שלך)
 ***************************************************/
var API_URL = "https://api.fireberry.com/api/record"; 
var QUERY_API_URL = "https://api.fireberry.com/api/query"; 



const FIREBERRY_API_KEY = PropertiesService.getScriptProperties().getProperty("FIREBERRY_API_KEY");
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

var MAX_RETRIES = 3;       
var RETRY_DELAY_MS = 3000;
// Add these with your other global variables
var cycleCache = {};
var guideCache = {};


/***************************************************
 * לוגים – ניתן לעיין בהם ב-Apps Script Logs
 ***************************************************/
function logMessage(msg) {
  Logger.log(msg);
}

/***************************************************
 * doGet – מחזיר את קובץ ה-HTML (Index)
 ***************************************************/
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index").evaluate();
}

/***************************************************
 * פונקציית getBranchById – שליפת אובייקט סניף (ObjectType=1009)
 ***************************************************/
function getBranchById(branchId) {
  if (!branchId) return null;
  var url = API_URL + "/1009/" + branchId;
  var data = sendGetWithRetry(url, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data) {
    Logger.log("Failed to fetch branch: " + branchId);
    return null;
  }
  return data.data.Record || null;
}

/***************************************************
 * מטמון לסניפים – branchCache
 ***************************************************/
var branchCache = {};

/***************************************************
 * getBranchName – מחזירה את שם הסניף לפי branchId (מתוך אובייקט 1009)
 ***************************************************/
function getBranchName(branchId) {
  if (!branchId) return "לא צויין";
  if (branchCache[branchId]) return branchCache[branchId];
  var branchObj = getBranchById(branchId);
  var name = branchObj && branchObj.name ? branchObj.name : "לא נמצא";
  branchCache[branchId] = name;
  return name;
}

/***************************************************
 * calcDailyForecastSS – חישוב יומי
 ***************************************************/
/***************************************************
calcDailyForecastSS – חישוב יומי עם שיפורים
***************************************************/
function calcDailyForecastSS(dayStr) {
  Logger.log("calcDailyForecastSS called: " + dayStr);
  var parts = dayStr.split("-");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var day = parseInt(parts[2], 10);
  var startDate = new Date(year, month - 1, day, 0, 0, 0);
  var endDate = new Date(year, month - 1, day + 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);
  var allMeetings = getMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total daily meetings: " + allMeetings.length);
  
  // השתמש בפונקציה החדשה לעיבוד פגישות במנות
  return processMeetingsInBatches(allMeetings);
}

/***************************************************
 * calcMonthlyForecastSS – חישוב חודשי
 ***************************************************/
/*function calcMonthlyForecastSS(monthStr) {
  Logger.log("calcMonthlyForecastSS called: " + monthStr);
  var parts = monthStr.split("-");
  var chosenYear = parseInt(parts[0], 10);
  var chosenMon = parseInt(parts[1], 10);
  var startDate = new Date(chosenYear, chosenMon - 1, 1, 0, 0, 0);
  var endDate = new Date(chosenYear, chosenMon, 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);
  
  // First try to get all meetings at once
  var allMeetings = getMeetingsForRange(startDateStr, endDateStr);
  
  // If we get too many meetings or encounter errors, process by week instead
  if (allMeetings.length > 5000 || allMeetings.length === 0) {
    Logger.log("Large dataset detected, processing by week");
    return processLargeMonthByWeeks(chosenYear, chosenMon);
  }
  
  Logger.log("Total monthly meetings: " + allMeetings.length);
  var totalRevenue = 0;
  var totalCost = 0;
  
  // Process in batches of 1000 meetings
  var batchSize = 1000;
  for (var i = 0; i < allMeetings.length; i += batchSize) {
    var endIndex = Math.min(i + batchSize, allMeetings.length);
    var batchResults = processMeetingBatch(allMeetings.slice(i, endIndex));
    totalRevenue += batchResults.revenue;
    totalCost += batchResults.cost;
    
    // Add a small delay between batch processing
    if (endIndex < allMeetings.length) {
      Utilities.sleep(500);
    }
  }
  
  return { success: true, totalRevenue: totalRevenue, totalCost: totalCost };
}*/

//**
 /* calcMonthlyForecastSS - חישוב תחזית חודשית עם סיווג לפי:
 * sivug = "cycle" | "lesson" | "branch"
 * @param {string} monthStr - פורמט "YYYY-MM", לדוגמה "2025-04"
 * @param {string} sivug - סוג הסיווג: "cycle" (לפי סוג המחזור),
 *                         "lesson" (לפי סוג ההדרכה), או "branch" (לפי סניף)
 * @return {object} אובייקט JSON המכיל totalRevenue, totalCost, 
 *                  ופירוט בהתאם לסיווג (כגון privateCount, frontaliCount, branchDetails וכו')
 */
function calcMonthlyForecastSS(monthStr, sivug) {
  Logger.log("calcMonthlyForecastSS called: " + monthStr + ", sivug: " + sivug);

  // פירוק החודש
  var parts = monthStr.split("-");
  var chosenYear = parseInt(parts[0], 10);
  var chosenMon = parseInt(parts[1], 10);

  // חישוב תאריכים: תחילת החודש ועד תחילת החודש הבא
  var startDate = new Date(chosenYear, chosenMon - 1, 1, 0, 0, 0);
  var endDate = new Date(chosenYear, chosenMon, 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);

  // שליפת כל הפגישות לחודש
  var allMeetings = getMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total monthly meetings: " + allMeetings.length);

  // עיבוד הפגישות לפי סוג הסיווג המבוקש
  var result = processMeetingsBySivug(allMeetings, sivug);

  // הבטחת ערכי ברירת מחדל
  result.totalRevenue = result.totalRevenue || 0;
  result.totalCost = result.totalCost || 0;
  result.success = true;
  
  Logger.log("Monthly Forecast Result: " + JSON.stringify(result));
  return result;
}

/**
 * calcMonthlyForecastNoStatusFilter – Computes the monthly forecast without
 * filtering out meetings based on status.
 *
 * @param {string} monthStr - Format "YYYY-MM", e.g. "2025-04"
 * @param {string} sivug - Grouping type: "cycle", "lesson", or "branch"
 * @return {object} JSON object with totalRevenue, totalCost, and details based on the grouping.
 */
function calcMonthlyForecastNoStatusFilter(monthStr, sivug) {
  Logger.log("calcMonthlyForecastNoStatusFilter called: " + monthStr + ", sivug: " + sivug);

  // Parse the chosen month
  var parts = monthStr.split("-");
  var chosenYear = parseInt(parts[0], 10);
  var chosenMon = parseInt(parts[1], 10);

  // Calculate the start and end dates for the month
  var startDate = new Date(chosenYear, chosenMon - 1, 1, 0, 0, 0);
  var endDate = new Date(chosenYear, chosenMon, 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);

  // Use the new function to get all meetings without filtering on status
  var allMeetings = getMeetingsForRangeWithoutStatusFilter(startDateStr, endDateStr);
  Logger.log("Total monthly meetings (no status filter): " + allMeetings.length);

  // Process the meetings by the chosen grouping (sivug)
  var result = processMeetingsBySivug(allMeetings, sivug);

  // Ensure default values are set
  result.totalRevenue = result.totalRevenue || 0;
  result.totalCost = result.totalCost || 0;
  result.success = true;
  
  Logger.log("Monthly Forecast (No Status Filter) Result: " + JSON.stringify(result));
  return result;
}

function getMeetingsForRangeWithoutStatusFilter(startDateStr, endDateStr) {
  Logger.log("getMeetingsForRangeWithoutStatusFilter: " + startDateStr + " to " + endDateStr);
  var allResults = [];
  var pageNumber = 1;
  var pageSize = 500; // Increased from 200
  var maxPages = 100;  // Increased from 10
  
  while (true) {
    Logger.log("Fetching meetings page=" + pageNumber);
    var queryPayload = {
      objecttype: 6,
      page_size: pageSize,
      page_number: pageNumber,
      fields: "activityid,scheduledstart,scheduledend,pcfsystemfield485,pcfsystemfield542,pcfsystemfield498,statuscode,pcfsystemfield559,pcfsystemfield545",
      query:
        "(scheduledstart >= " + startDateStr + ") AND " +
        "(scheduledstart < " + endDateStr + ") AND " +
        "(pcfsystemfield498 is-not-null)"
    };

    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) {
      Logger.log("Error or null returned for page " + pageNumber + ", stopping.");
      break;
    }
    
    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    Logger.log("Got " + chunk.length + " meetings in this page.");
    allResults = allResults.concat(chunk);
    
    if (chunk.length < pageSize || pageNumber >= maxPages) {
      Logger.log("Less than pageSize or reached maxPages => stopping pagination.");
      break;
    }
    
    pageNumber++;
    Utilities.sleep(300);
  }
  
  Logger.log("Total meetings after pagination: " + allResults.length);
  return allResults;
}


/**
 * processMeetingsBySivug - מעבדת מערך פגישות ומציבה את הסיכומים הכוללים
 * וגם פירוט לפי הסיווג:
 *  - "cycle": לפי סוג המחזור (pcfsystemfield549: 1=פרטי, 2=מוסדי, 3=מוסדי פר ילד)
 *  - "lesson": לפי סוג ההדרכה (pcfsystemfield542, כשהערכים הם טקסט: 
 *              "שיעור פרונטאלי", "שיעור פרטי", "שיעור אונליין קבוצתי", "תמיכה או בניית חומרים")
 *  - "branch": לפי הסניף (pcfsystemfield451, כאשר getBranchName מחזירה את השם)
 * @param {array} meetings - מערך הפגישות
 * @param {string} sivug - סוג הסיווג
 * @return {object} אובייקט עם totalRevenue, totalCost, ופירוט בהתאם לסיווג
 */
function processMeetingsBySivug(meetings, sivug) {
  var totalRevenue = 0;
  var totalCost = 0;
  
  // Variables for grouping by cycle
  var privateCount = 0, privateSum = 0;
  var mosediCount = 0, mosediSum = 0;
  var mosediChildCount = 0, mosediChildSum = 0;
  
  // Variables for grouping by lesson
  var frontaliCount = 0, frontaliSum = 0;
  var privateLessonCount = 0, privateLessonSum = 0;
  var onlineCount = 0, onlineSum = 0;
  var supportCount = 0, supportSum = 0;
  
  // For grouping by branch
  var branchMap = {};
  
  // Process meetings in batches
  var batchSize = 1000;
  for (var i = 0; i < meetings.length; i += batchSize) {
    var endIndex = Math.min(i + batchSize, meetings.length);
    var batchMeetings = meetings.slice(i, endIndex);
    
    for (var j = 0; j < batchMeetings.length; j++) {
      var meeting = batchMeetings[j];
      var cycleId = meeting.pcfsystemfield498;
      if (!cycleId) continue;
      var cycle = getCycleById(cycleId);
      if (!cycle) continue;
      
      // Initialize revenue and cost variables
      var inc = 0;
      var cost = 0;
      
      // If the meeting has a status and it is "התקיימה", then use the actual values from the meeting
      if (meeting.statuscode) {
        // If there is a status, check if it's "התקיימה"
        if (meeting.statuscode.trim() === "התקיימה") {
          inc = parseFloat(meeting.pcfsystemfield559 || 0);
          cost = parseFloat(meeting.pcfsystemfield545 || 0);
        } else {
          // For any other status, skip this meeting (do not include it)
          continue; // (if within a loop) or simply skip adding its values.
        }
      } else {
        // No status provided => use forecast logic
        inc = getIncomePerMeeting(cycle);
        cost = calculateMeetingCost(meeting);
      }
      
      totalRevenue += inc;
      totalCost += cost;
      
      // Group the results according to the requested grouping type
      if (sivug === "cycle") {
        var cType = parseInt(cycle.pcfsystemfield549 || "0", 10);
        if (cType === 1) {
          privateCount++;
          privateSum += inc;
        } else if (cType === 2) {
          mosediCount++;
          mosediSum += inc;
        } else if (cType === 3) {
          mosediChildCount++;
          mosediChildSum += inc;
        }
      } else if (sivug === "lesson") {
        var lessonTypeText = (meeting.pcfsystemfield542 || "").trim();
        switch (lessonTypeText) {
          case "שיעור פרונטאלי":
            frontaliCount++;
            frontaliSum += inc;
            break;
          case "שיעור פרטי":
            privateLessonCount++;
            privateLessonSum += inc;
            break;
          case "שיעור אונליין קבוצתי":
            onlineCount++;
            onlineSum += inc;
            break;
          case "תמיכה או בניית חומרים":
            supportCount++;
            supportSum += inc;
            break;
          default:
            Logger.log("Unknown lessonTypeText: " + lessonTypeText);
            break;
        }
      } else if (sivug === "branch") {
        var branchId = cycle.pcfsystemfield451;
        var branchName = getBranchName(branchId);
        if (!branchMap[branchName]) {
          branchMap[branchName] = { count: 0, sumRevenue: 0 };
        }
        branchMap[branchName].count++;
        branchMap[branchName].sumRevenue += inc;
      }
    }
    
    if (endIndex < meetings.length) {
      Utilities.sleep(500);
    }
  }
  
  var result = {
    success: true,
    totalRevenue: totalRevenue,
    totalCost: totalCost
  };
  
  if (sivug === "cycle") {
    result.privateCount = privateCount;
    result.privateSum = privateSum;
    result.mosediCount = mosediCount;
    result.mosediSum = mosediSum;
    result.mosediChildCount = mosediChildCount;
    result.mosediChildSum = mosediChildSum;
  } else if (sivug === "lesson") {
    result.frontaliCount = frontaliCount;
    result.frontaliSum = frontaliSum;
    result.privateLessonCount = privateLessonCount;
    result.privateLessonSum = privateLessonSum;
    result.onlineCount = onlineCount;
    result.onlineSum = onlineSum;
    result.supportCount = supportCount;
    result.supportSum = supportSum;
  } else if (sivug === "branch") {
    var branchDetails = [];
    for (var bName in branchMap) {
      branchDetails.push({
        branchName: bName,
        count: branchMap[bName].count,
        sumRevenue: branchMap[bName].sumRevenue
      });
    }
    result.branchDetails = branchDetails;
  }
  
  return result;
}

/*function processMeetingsBySivug(meetings, sivug) {
  var totalRevenue = 0;
  var totalCost = 0;
  
  if (sivug === "cycle") {
    // New logic for cycle grouping:
    // Group meetings by cycle ID.
    var cyclesMap = {};
    meetings.forEach(function(meeting) {
      var cycleId = meeting.pcfsystemfield498;
      if (!cycleId) return;
      if (!cyclesMap[cycleId]) {
        cyclesMap[cycleId] = [];
      }
      cyclesMap[cycleId].push(meeting);
    });
    
    // Process each cycle separately.
    for (var cycleId in cyclesMap) {
      var cycleMeetings = cyclesMap[cycleId];
      var cycle = getCycleById(cycleId);
      if (!cycle) continue;
      
      // The maximum number of forecast meetings is given in pcfsystemfield233
      var maxForecast = parseInt(cycle.pcfsystemfield233 || "0", 10);
      
      // Filter meetings to those with no status (i.e. forecast meetings)
      var forecastMeetings = cycleMeetings.filter(function(m) {
        // If the meeting has a status (non-empty) then we do NOT use forecast logic.
        return !(m.statuscode && m.statuscode.trim() !== "");
      });
      
      var cycleRevenue = 0;
      var cycleCost = 0;
      var numForecast = forecastMeetings.length;
      
      // Process forecast meetings up to maxForecast.
      for (var i = 0; i < Math.min(numForecast, maxForecast); i++) {
        var meeting = forecastMeetings[i];
        // Use forecast logic for each meeting
        var inc = getIncomePerMeeting(cycle);
        var cost = calculateMeetingCost(meeting);
        cycleRevenue += inc;
        cycleCost += cost;
      }
      
      // If there are fewer forecast meetings than maxForecast, pad using the last forecast meeting's values.
      if (numForecast > 0 && numForecast < maxForecast) {
        var missing = maxForecast - numForecast;
        var lastMeeting = forecastMeetings[numForecast - 1];
        var inc = getIncomePerMeeting(cycle);
        var cost = calculateMeetingCost(lastMeeting);
        cycleRevenue += missing * inc;
        cycleCost += missing * cost;
      }
      
      totalRevenue += cycleRevenue;
      totalCost += cycleCost;
    }
    
  } else {
    // For non-cycle grouping (i.e. for "lesson" and "branch") use the original per‐meeting logic.
    var privateCount = 0, privateSum = 0;
    var mosediCount = 0, mosediSum = 0;
    var mosediChildCount = 0, mosediChildSum = 0;
    var frontaliCount = 0, frontaliSum = 0;
    var privateLessonCount = 0, privateLessonSum = 0;
    var onlineCount = 0, onlineSum = 0;
    var supportCount = 0, supportSum = 0;
    var branchMap = {};
    
    var batchSize = 1000;
    for (var i = 0; i < meetings.length; i += batchSize) {
      var endIndex = Math.min(i + batchSize, meetings.length);
      var batchMeetings = meetings.slice(i, endIndex);
      
      for (var j = 0; j < batchMeetings.length; j++) {
        var meeting = batchMeetings[j];
        var cycleId = meeting.pcfsystemfield498;
        if (!cycleId) continue;
        var cycle = getCycleById(cycleId);
        if (!cycle) continue;
        
        // If the meeting has a status and it is "התקיימה" then use actual values;
        // Otherwise (if status is empty) use forecast logic.
        var inc = 0, cost = 0;
        if (meeting.statuscode && meeting.statuscode.trim() === "התקיימה") {
          inc = parseFloat(meeting.pcfsystemfield559 || 0);
          cost = parseFloat(meeting.pcfsystemfield545 || 0);
        } else {
          inc = getIncomePerMeeting(cycle);
          cost = calculateMeetingCost(meeting);
        }
        totalRevenue += inc;
        totalCost += cost;
        
        if (sivug === "lesson") {
          var lessonTypeText = (meeting.pcfsystemfield542 || "").trim();
          switch (lessonTypeText) {
            case "שיעור פרונטאלי":
              frontaliCount++;
              frontaliSum += inc;
              break;
            case "שיעור פרטי":
              privateLessonCount++;
              privateLessonSum += inc;
              break;
            case "שיעור אונליין קבוצתי":
              onlineCount++;
              onlineSum += inc;
              break;
            case "תמיכה או בניית חומרים":
              supportCount++;
              supportSum += inc;
              break;
            default:
              Logger.log("Unknown lessonTypeText: " + lessonTypeText);
              break;
          }
        } else if (sivug === "branch") {
          var branchId = cycle.pcfsystemfield451;
          var branchName = getBranchName(branchId);
          if (!branchMap[branchName]) {
            branchMap[branchName] = { count: 0, sumRevenue: 0 };
          }
          branchMap[branchName].count++;
          branchMap[branchName].sumRevenue += inc;
        }
      }
      
      if (endIndex < meetings.length) {
        Utilities.sleep(500);
      }
    }
    
    // For lesson or branch grouping, attach breakdown details.
    if (sivug === "lesson") {
      var result = {
        success: true,
        frontaliCount: frontaliCount,
        frontaliSum: frontaliSum,
        privateLessonCount: privateLessonCount,
        privateLessonSum: privateLessonSum,
        onlineCount: onlineCount,
        onlineSum: onlineSum,
        supportCount: supportCount,
        supportSum: supportSum
      };
    } else if (sivug === "branch") {
      var branchDetails = [];
      for (var bName in branchMap) {
        branchDetails.push({
          branchName: bName,
          count: branchMap[bName].count,
          sumRevenue: branchMap[bName].sumRevenue
        });
      }
      var result = {
        success: true,
        branchDetails: branchDetails
      };
    }
  }
  
  // Return overall totals for revenue and cost.
  return { success: true, totalRevenue: totalRevenue, totalCost: totalCost };
}*/


function processLargeMonthByWeeks(year, month) {
  var daysInMonth = getDaysInMonth(year, month);
  var totalRevenue = 0;
  var totalCost = 0;
  
  // Process one week at a time
  for (var startDay = 1; startDay <= daysInMonth; startDay += 7) {
    var endDay = Math.min(startDay + 6, daysInMonth);
    
    var startDate = new Date(year, month - 1, startDay, 0, 0, 0);
    var endDate = new Date(year, month - 1, endDay + 1, 0, 0, 0);
    var startDateStr = toIsoString(startDate);
    var endDateStr = toIsoString(endDate);
    
    Logger.log("Processing week from " + startDay + " to " + endDay);
    var weekMeetings = getMeetingsForRange(startDateStr, endDateStr);
    
    var weekResults = processMeetingBatch(weekMeetings);
    totalRevenue += weekResults.revenue;
    totalCost += weekResults.cost;
    
    // Add delay between weeks
    Utilities.sleep(1000);
  }
  
  return { success: true, totalRevenue: totalRevenue, totalCost: totalCost };
}

/***************************************************
 * calcWeeklyForecastSS – חישוב שבועי עם אפשרות סיווג לפי:
 * sivug = "cycle" | "lesson" | "branch"
 ***************************************************/
/***************************************************
calcWeeklyForecastSS – חישוב שבועי משופר
***************************************************/
function calcWeeklyForecastSS(monthStr, weekStr, sivug) {
  Logger.log("calcWeeklyForecastSS called: month=" + monthStr + ", week=" + weekStr + ", sivug=" + sivug);
  var chosenWeek = parseInt(weekStr, 10);
  var parts = monthStr.split("-");
  var chosenYear = parseInt(parts[0], 10);
  var chosenMon = parseInt(parts[1], 10);

  var startDay = (chosenWeek - 1)*7 + 1;
  var lastDayInMonth = getDaysInMonth(chosenYear, chosenMon);
  var endDay = startDay + 6;
  if (endDay > lastDayInMonth) endDay = lastDayInMonth;

  var startDate = new Date(chosenYear, chosenMon - 1, startDay, 0, 0, 0);
  var endDate = new Date(chosenYear, chosenMon - 1, endDay + 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);

  var allMeetings = getMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total weekly meetings found: " + allMeetings.length);

  // יצירת אובייקט תוצאה עם חלוקה לפי סיווג
  var result = processWeeklyMeetingsBySivug(allMeetings, sivug);
  return result;
}

/***************************************************
processWeeklyMeetingsBySivug – עיבוד פגישות שבועיות לפי סיווג
***************************************************/
function processWeeklyMeetingsBySivug(meetings, sivug) {
  var totalRevenue = 0;
  var totalCost = 0;

  // לסיווג לפי "cycle" (מחזור) – pcfsystemfield549
  var privateCount = 0, privateSum = 0;
  var mosediCount = 0, mosediSum = 0;
  var mosediChildCount = 0, mosediChildSum = 0;

  // לסיווג לפי "lesson" – pcfsystemfield542
  var frontaliCount = 0, frontaliSum = 0;
  var privateLessonCount = 0, privateLessonSum = 0;
  var onlineCount = 0, onlineSum = 0;
  var supportCount = 0, supportSum = 0;

  // לסיווג לפי "branch" – pcfsystemfield451
  var branchMap = {};

  // עיבוד במנות של 1000 פגישות
  var batchSize = 1000;
  for (var i = 0; i < meetings.length; i += batchSize) {
    var endIndex = Math.min(i + batchSize, meetings.length);
    var batchMeetings = meetings.slice(i, endIndex);
    
    for (var j = 0; j < batchMeetings.length; j++) {
      var meeting = batchMeetings[j];
      var cycleId = meeting.pcfsystemfield498;
      if (!cycleId) continue;
      var cycle = getCycleById(cycleId);
      if (!cycle) continue;
      var inc = getIncomePerMeeting(cycle);
      var cost = calculateMeetingCost(meeting);
      totalRevenue += inc;
      totalCost += cost;

      if (sivug === "cycle") {
        var cType = parseInt(cycle.pcfsystemfield549 || "0", 10);
        if (cType === 1) {
          privateCount++;
          privateSum += inc;
        } else if (cType === 2) {
          mosediCount++;
          mosediSum += inc;
        } else if (cType === 3) {
          mosediChildCount++;
          mosediChildSum += inc;
        }
      } else if (sivug === "lesson") {
        var lessonTypeText = meeting.pcfsystemfield542 || "";
        switch (lessonTypeText) {
          case "שיעור פרונטאלי":
            frontaliCount++;
            frontaliSum += inc;
            break;
          case "שיעור פרטי":
            privateLessonCount++;
            privateLessonSum += inc;
            break;
          case "שיעור אונליין קבוצתי":
            onlineCount++;
            onlineSum += inc;
            break;
          case "תמיכה או בניית חומרים":
            supportCount++;
            supportSum += inc;
            break;
          default:
            Logger.log("lessonTypeText לא מוכר: " + lessonTypeText);
            break;
        }
      } else if (sivug === "branch") {
        var branchId = cycle.pcfsystemfield451;
        var branchName = getBranchName(branchId);
        if (!branchMap[branchName]) {
          branchMap[branchName] = { count: 0, sumRevenue: 0 };
        }
        branchMap[branchName].count++;
        branchMap[branchName].sumRevenue += inc;
      }
    }
    
    // הוספת השהייה קטנה בין עיבוד מנות
    if (endIndex < meetings.length) {
      Utilities.sleep(500);
    }
  }

  var result = {
    success: true,
    totalRevenue: totalRevenue,
    totalCost: totalCost
  };

  if (sivug === "cycle") {
    result.privateCount = privateCount;
    result.privateSum = privateSum;
    result.mosediCount = mosediCount;
    result.mosediSum = mosediSum;
    result.mosediChildCount = mosediChildCount;
    result.mosediChildSum = mosediChildSum;
  } else if (sivug === "lesson") {
    result.frontaliCount = frontaliCount;
    result.frontaliSum = frontaliSum;
    result.privateLessonCount = privateLessonCount;
    result.privateLessonSum = privateLessonSum;
    result.onlineCount = onlineCount;
    result.onlineSum = onlineSum;
    result.supportCount = supportCount;
    result.supportSum = supportSum;
  } else if (sivug === "branch") {
    var branchDetails = [];
    for (var bName in branchMap) {
      branchDetails.push({
        branchName: bName,
        count: branchMap[bName].count,
        sumRevenue: branchMap[bName].sumRevenue
      });
    }
    result.branchDetails = branchDetails;
  }

  return result;
}


/***************************************************
 * calcDailyForecastSS – חישוב יומי
 ***************************************************/
function calcDailyForecastSS(dayStr) {
  Logger.log("calcDailyForecastSS called: " + dayStr);
  var parts = dayStr.split("-");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  var day = parseInt(parts[2], 10);
  var startDate = new Date(year, month - 1, day, 0, 0, 0);
  var endDate = new Date(year, month - 1, day + 1, 0, 0, 0);
  var startDateStr = toIsoString(startDate);
  var endDateStr = toIsoString(endDate);
  var allMeetings = getMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total daily meetings: " + allMeetings.length);
  var totalRevenue = 0;
  var totalCost = 0;
  for (var i = 0; i < allMeetings.length; i++) {
    var meeting = allMeetings[i];
    var cycleId = meeting.pcfsystemfield498;
    if (!cycleId) continue;
    var cycle = getCycleById(cycleId);
    if (!cycle) continue;
    var inc = getIncomePerMeeting(cycle);
    var cost = calculateMeetingCost(meeting);
    totalRevenue += inc;
    totalCost += cost;
  }
  return { success: true, totalRevenue: totalRevenue, totalCost: totalCost };
}

/***************************************************
 * getMeetingsForRange – שליפת פגישות בטווח [startDateStr, endDateStr)
 * כולל Pagination (pageSize=200)
 ***************************************************/
function getMeetingsForRange(startDateStr, endDateStr) {
  Logger.log("getMeetingsForRange: " + startDateStr + " to " + endDateStr);
  var allResults = [];
  var pageNumber = 1;
  var pageSize = 500; // Increased from 200
  var maxPages = 100;  // Increased from 10
  
  while (true) {
    Logger.log("Fetching meetings page=" + pageNumber);
    var queryPayload = {
      objecttype: 6,
      page_size: pageSize,
      page_number: pageNumber,
      fields: "activityid,scheduledstart,scheduledend,pcfsystemfield485,pcfsystemfield542,pcfsystemfield498,statuscode",
      query:
        "(scheduledstart >= " + startDateStr + ") AND " +
        "(scheduledstart < " + endDateStr + ") AND " +
        "(pcfsystemfield498 is-not-null) AND " +
        "(statuscode is-null)"
    };
    
    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) {
      Logger.log("Error or null returned for page " + pageNumber + ", stopping.");
      break;
    }
    
    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    Logger.log("Got " + chunk.length + " meetings in this page.");
    allResults = allResults.concat(chunk);
    
    if (chunk.length < pageSize) {
      Logger.log("Less than pageSize => no more pages.");
      break;
    }
    
    if (pageNumber >= maxPages) {
      Logger.log("Reached maxPages=" + maxPages + ", stopping pagination.");
      break;
    }
    
    pageNumber++;
    
    // Add a small delay between pagination requests to avoid rate limiting
    Utilities.sleep(300);
  }
  
  Logger.log("Total meetings after pagination: " + allResults.length);
  return allResults;
}

/***************************************************
 * getCycleById – שליפת מחזור (ObjectType=1000) לפי cycleId
 ***************************************************/
// Update getCycleById to use cache
function getCycleById(cycleId) {
  if (!cycleId) return null;
  if (cycleCache[cycleId]) return cycleCache[cycleId];
  
  var url = API_URL + "/1000/" + cycleId;
  var data = sendGetWithRetry(url, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data) {
    Logger.log("Failed to fetch cycle: " + cycleId);
    return null;
  }
  
  var cycle = data.data.Record || null;
  if (cycle) cycleCache[cycleId] = cycle;
  return cycle;
}

/***************************************************
 * getIncomePerMeeting – חישוב הכנסה לפגישה לפי סוג המחזור
 * pcfsystemfield549: 1=פרטי, 2=מוסדי, 3=מוסדי פר ילד
 ***************************************************/
function getIncomePerMeeting(cycle) {
  var cycleType = parseInt(cycle.pcfsystemfield549 || "0", 10);
  var numMeetings = parseInt(cycle.pcfsystemfield88 || "0", 10);
  switch (cycleType) {
    case 1: // פרטי
      if (numMeetings === 0) return 0;
      var totalPrivateIncome = getSumOfRegistrations(cycle.customobject1000id);
      var netAfterVAT = totalPrivateIncome / 1.18;
      return netAfterVAT / numMeetings;
    case 2: // מוסדי
      var totalMosedi = parseFloat(cycle.pcfsystemfield550 || 0);
      return totalMosedi / 1.18;
    case 3: // מוסדי פר ילד
      var kidsCount = parseFloat(cycle.pcfsystemfield192 || 0);
      if (kidsCount === 0) {
        kidsCount = parseFloat(cycle.pcfsystemfield552 || 0);
      }
      var pricePerKid = parseFloat(cycle.pcfsystemfield551 || 0);
      return kidsCount * pricePerKid;
    default:
      return 0;
  }
}

/***************************************************
 * getSumOfRegistrations – שליפת סכום הרשמות (ObjectType=33)
 ***************************************************/
var registrationCache = {};

function getSumOfRegistrations(cycleId) {
  // Check cache first
  if (registrationCache[cycleId] !== undefined) {
    return registrationCache[cycleId];
  }
  
  var queryPayload = {
    objecttype: 33,
    page_size: 200,
    page_number: 1,
    fields: "accountproductid,pcfsystemfield289",
    query: "(pcfsystemfield53 = " + cycleId + ")"
  };
  
  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data) return 0;
  
  var regs = (data.data && data.data.Data) ? data.data.Data : [];
  var sum = 0;
  for (var i = 0; i < regs.length; i++) {
    sum += parseFloat(regs[i].pcfsystemfield289 || 0);
  }
  
  // Store in cache
  registrationCache[cycleId] = sum;
  return sum;
}

/***************************************************
 * calculateMeetingCost – חישוב עלות פגישה למדריך
 ***************************************************/
function calculateMeetingCost(meeting) {
  var guideId = meeting.pcfsystemfield485;
  // במקום parseInt, נקבל את הטקסט (שיעור פרונטאלי, שיעור פרטי, וכו')
  var lessonTypeText = meeting.pcfsystemfield542 || "";

  // נקבל את תעריף המדריך לפי הטקסט (בפונקציה getGuideRate המעודכנת)
  var rate = getGuideRate(guideId, lessonTypeText);
 
  var start = new Date(meeting.scheduledstart);
  var end = new Date(meeting.scheduledend);
  var diffMs = end - start;
  var diffMin = diffMs / (1000 * 60);
  if (diffMin < 0) diffMin = 0;
  var cost = (diffMin / 60) * rate;
  var freelancer = isGuideFreelancer(guideId);
  if (!freelancer) {
    cost *= 1.25;
  }
  return cost;
}

/***************************************************
 * getGuideRate – תעריף לשעה לפי סוג ההדרכה (lessonType)
 ***************************************************/
/**
 * getGuideRate(guideId, lessonTypeText)
 * lessonTypeText יכול להיות:
 *  "שיעור פרונטאלי"
 *  "שיעור פרטי"
 *  "שיעור אונליין קבוצתי"
 *  "תמיכה או בניית חומרים"
 *  "מותאם אישית" (או ערך אחר)
 */
function getGuideRate(guideId, lessonTypeText) {
  var guide = getGuideById(guideId);
  if (!guide) return 0;

  switch (lessonTypeText) {
    case "שיעור פרונטאלי":
      return parseFloat(guide.pcfsystemfield536 || 0);
    case "שיעור פרטי":
      return parseFloat(guide.pcfsystemfield549 || 0);
    case "שיעור אונליין קבוצתי":
      return parseFloat(guide.pcfsystemfield551 || 0);
    case "תמיכה או בניית חומרים":
      return parseFloat(guide.pcfsystemfield543 || 0);
    case "מותאם אישית":
      // אם יש שדה מיוחד למותאם אישית, אפשר להוסיף כאן
      return 0;
    default:
      // אם יש ערך טקסט אחר לא צפוי
      Logger.log("lessonTypeText לא מוכר: " + lessonTypeText);
      return 0;
  }
}


/***************************************************
 * isGuideFreelancer – האם המדריך פרילנסר (pcfsystemfield563=1)
 ***************************************************/
function isGuideFreelancer(guideId) {
  var guide = getGuideById(guideId);
  if (!guide) return false;
  return (+guide.pcfsystemfield563 === 1);
}


/***************************************************
 * getGuideById – שליפת מדריך (ObjectType=1002)
 ***************************************************/
// Update getGuideById to use cache
function getGuideById(guideId) {
  if (!guideId) return null;
  if (guideCache[guideId]) return guideCache[guideId];
  
  var url = API_URL + "/1002/" + guideId;
  var data = sendGetWithRetry(url, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data) {
    Logger.log("Failed to get guide: " + guideId);
    return null;
  }
  
  var guide = data.data.Record || null;
  if (guide) guideCache[guideId] = guide;
  return guide;
}
/***************************************************
 * sendRequestWithRetry – בקשת POST עם retry
 ***************************************************/
function sendRequestWithRetry(url, payload, maxRetries, initialRetryDelayMs) {
  var retryDelayMs = initialRetryDelayMs;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log("POST attempt " + attempt + " to " + url);
      var options = {
        method: 'POST',
        contentType: 'application/json',
        headers: {
          tokenid: FIREBERRY_API_KEY,
          accept: 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      var response = UrlFetchApp.fetch(url, options);
      var httpStatus = response.getResponseCode();
      var responseText = response.getContentText();
      
      if (httpStatus === 200) {
        Logger.log("POST success on attempt " + attempt);
        return JSON.parse(responseText);
      }
      
      if (httpStatus === 429) {
        Logger.log("Got 429 on POST attempt " + attempt);
        if (attempt < maxRetries) {
          // Exponential backoff: double the delay with each retry
          Utilities.sleep(retryDelayMs);
          retryDelayMs *= 2;
          continue;
        } else {
          Logger.log("Exceeded max retries for 429. Last response: " + responseText);
          return null;
        }
      }
      
      Logger.log("HTTP " + httpStatus + " Error: " + responseText);
      return null;
    } catch (err) {
      Logger.log("Exception in POST attempt " + attempt + ": " + err);
      if (attempt < maxRetries) {
        Utilities.sleep(retryDelayMs);
        retryDelayMs *= 2;
      } else {
        return null;
      }
    }
  }
  return null;
}
/***************************************************
 * sendGetWithRetry – בקשת GET עם retry
 ***************************************************/
function sendGetWithRetry(url, maxRetries, retryDelayMs) {
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log("GET attempt " + attempt + " => " + url);
      var options = {
        method: 'GET',
        headers: {
          tokenid: FIREBERRY_API_KEY,
          accept: 'application/json'
        },
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(url, options);
      var httpStatus = response.getResponseCode();
      var responseText = response.getContentText();
      if (httpStatus === 200) {
        Logger.log("GET success on attempt " + attempt);
        return JSON.parse(responseText);
      }
      if (httpStatus === 429) {
        Logger.log("Got 429 on GET attempt " + attempt);
        if (attempt < maxRetries) {
          Utilities.sleep(retryDelayMs);
          continue;
        } else {
          Logger.log("Exceeded max retries for 429. Last response: " + responseText);
          return null;
        }
      }
      Logger.log("HTTP " + httpStatus + " Error (GET): " + responseText);
      return null;
    } catch (err) {
      Logger.log("Exception in GET attempt " + attempt + ": " + err);
      if (attempt < maxRetries) {
        Utilities.sleep(retryDelayMs);
      } else {
        return null;
      }
    }
  }
  return null;
}

/***************************************************
 * toIsoString – המרת Date לפורמט ISO (2025-03-06T10:00:00)
 ***************************************************/
function toIsoString(d) {
  var year = d.getFullYear();
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  var hour = ('0' + d.getHours()).slice(-2);
  var min = ('0' + d.getMinutes()).slice(-2);
  var sec = ('0' + d.getSeconds()).slice(-2);
  return year + "-" + month + "-" + day + "T" + hour + ":" + min + ":" + sec;
}

/***************************************************
 * getDaysInMonth – כמות הימים בחודש נתון
 ***************************************************/
function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

/***************************************************
clearAllCaches – ניקוי כל המטמונים
***************************************************/
function clearAllCaches() {
  branchCache = {};
  cycleCache = {};
  guideCache = {};
  registrationCache = {};
  Logger.log("All caches cleared");
}

/**
 * exportForecastToNewSheet – מייצאת את תוצאת התחזית לדוח בגוגל שיטס חדש.
 * הקובץ נוצר מהקוד ולא מתווסף לגליון קיים.
 *
 * @param {object} result - אובייקט התחזית, לדוגמה כפי שמוחזר מ־calcWeeklyForecastSS, calcMonthlyForecastSS או calcDailyForecastSS.
 * @param {string} sheetName - שם הקובץ לדוגמה "Forecast Report April 2025"
 * @return {string} ה-URL של הגיליון החדש.
 */
function exportForecastToNewSheet(result, sheetName) {
  try {
    // צור קובץ חדש עם השם המבוקש
    var ss = SpreadsheetApp.create(sheetName);
    var sheet = ss.getActiveSheet();
    sheet.clear();
    
    // בנה מערך של שורות לכתיבה
    var rows = [];
    
    // הוסף כותרת ראשית
    rows.push(["דוח תחזית", sheetName]);
    // במקום להוסיף שורות ריקות כ[]
    rows.push(["", ""]);
    
    // הוסף שורות עם סך ההכנסה וההוצאה
    rows.push(["סך הכנסה", result.totalRevenue]);
    rows.push(["סך הוצאה", result.totalCost]);
    rows.push(["", ""]);
    
    // הוסף פירוט בהתאם לסיווג
    if (result.privateCount !== undefined) {  // sivug = "cycle"
      rows.push(["סיווג לפי סוג מחזור"]);
      rows.push(["מחזור פרטי", result.privateCount, result.privateSum]);
      rows.push(["מחזור מוסדי", result.mosediCount, result.mosediSum]);
      rows.push(["מוסדי פר ילד", result.mosediChildCount, result.mosediChildSum]);
      rows.push(["", ""]);
    }
    
    if (result.frontaliCount !== undefined) {  // sivug = "lesson"
      rows.push(["סיווג לפי סוג ההדרכה"]);
      rows.push(["שיעור פרונטאלי", result.frontaliCount, result.frontaliSum]);
      rows.push(["שיעור פרטי", result.privateLessonCount, result.privateLessonSum]);
      rows.push(["אונליין קבוצתי", result.onlineCount, result.onlineSum]);
      rows.push(["תמיכה", result.supportCount, result.supportSum]);
      rows.push(["", ""]);
    }
    
    if (result.branchDetails && result.branchDetails.length > 0) {  // sivug = "branch"
      rows.push(["סיווג לפי סניף"]);
      rows.push(["שם הסניף", "מספר מפגשים", "סך הכנסה"]);
      for (var i = 0; i < result.branchDetails.length; i++) {
        var bd = result.branchDetails[i];
        rows.push([bd.branchName, bd.count, bd.sumRevenue]);
      }
      rows.push(["", ""]);
    }
    
    // חשב את מספר העמודות המקסימלי בכל השורות
    var maxCols = 0;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].length > maxCols) {
        maxCols = rows[i].length;
      }
    }
    
    // מלא כל שורה בערכים ריקים עד שתגיע לאורך maxCols
    for (var i = 0; i < rows.length; i++) {
      while (rows[i].length < maxCols) {
        rows[i].push("");
      }
    }
    
    // כתיבה לגיליון
    sheet.getRange(1, 1, rows.length, maxCols).setValues(rows);
    sheet.autoResizeColumns(1, maxCols);
    
    var url = ss.getUrl();
    Logger.log("Forecast exported to new spreadsheet: " + url);
    return url;
    
  } catch (e) {
    Logger.log("Error in exportForecastToNewSheet: " + e);
    throw e;
  }
}

/**
 * compareMonthlyForecastByBranch – משווה בין שני חודשים (לפי סניף)
 * עבור כל חודש, היא קוראת לפונקציה calcMonthlyForecastSS עם sivug = "branch"
 * ומחזירה טבלה שמשווה את ההכנסה בכל סניף בין שני החודשים.
 *
 * @param {string} month1 - פורמט "YYYY-MM", לדוגמה "2025-04"
 * @param {string} month2 - פורמט "YYYY-MM", לדוגמה "2025-05"
 * @return {object} אובייקט המכיל:
 *   - success: true אם הצליח,
 *   - month1, month2: החודשים שביקשת,
 *   - comparison: מערך שבו כל שורה היא [branchName, incomeMonth1, incomeMonth2, difference, percentageChange]
 */
function compareMonthlyForecastByBranch(month1, month2) {
  // Fetch forecast data for each month using your existing calcMonthlyForecastSS
  var result1 = calcMonthlyForecastNoStatusFilter(month1, "branch");
  var result2 = calcMonthlyForecastNoStatusFilter(month2, "branch");
  
  if (!result1 || !result1.success || !result2 || !result2.success) {
    return { success: false, error: "בעיה בשליפת נתוני התחזית לחודשים." };
  }
  
  // Build a branch map from the first month's details
  var branchMap = {};
  if (result1.branchDetails && result1.branchDetails.length > 0) {
    result1.branchDetails.forEach(function(item) {
      branchMap[item.branchName] = { incomeMonth1: item.sumRevenue, count1: item.count };
    });
  }
  
  // Update the map with the second month's data
  if (result2.branchDetails && result2.branchDetails.length > 0) {
    result2.branchDetails.forEach(function(item) {
      if (!branchMap[item.branchName]) {
        branchMap[item.branchName] = { incomeMonth1: 0, count1: 0 };
      }
      branchMap[item.branchName].incomeMonth2 = item.sumRevenue;
      branchMap[item.branchName].count2 = item.count;
    });
  }
  
  // Build the comparison array and accumulate totals
  var comparison = [];
  var totalMonth1 = 0, totalMonth2 = 0;
  for (var branchName in branchMap) {
    var inc1 = branchMap[branchName].incomeMonth1 || 0;
    var inc2 = branchMap[branchName].incomeMonth2 || 0;
    var diff = inc2 - inc1;
    var percentage = inc1 === 0 ? 0 : (diff / inc1 * 100);
    totalMonth1 += inc1;
    totalMonth2 += inc2;
    comparison.push([branchName, inc1, inc2, diff, percentage]);
  }
  // Add summary row at the end
  var totalDiff = totalMonth2 - totalMonth1;
  var totalPerc = totalMonth1 === 0 ? 0 : (totalDiff / totalMonth1 * 100);
  comparison.push(["סיכום", totalMonth1, totalMonth2, totalDiff, totalPerc]);
  
  return {
    success: true,
    month1: month1,
    month2: month2,
    comparison: comparison
  };
}


/**
 * exportCompareToSheet – מייצאת את תוצאת השוואת החודשים לדוח בגוגל שיטס חדש.
 * @param {object} compareResult - אובייקט ההשוואה כפי שמוחזר מ־compareMonthlyForecastByBranch.
 * @param {string} reportName - שם הדוח הרצוי.
 * @return {string} ה-URL של הגיליון החדש.
 */
function exportCompareToSheet(compareResult, reportName) {
  try {
    var ss = SpreadsheetApp.create(reportName);
    var sheet = ss.getActiveSheet();
    sheet.clear();
    
    var rows = [];
    rows.push(["השוואת חודשים לפי סניף", reportName]);
    rows.push([]);
    rows.push(["חודש ראשון", compareResult.month1, "חודש שני", compareResult.month2]);
    rows.push([]);
    rows.push(["סניף", "הכנסה (" + compareResult.month1 + ")", "הכנסה (" + compareResult.month2 + ")", "הפרש", "% שינוי"]);
    
    if (compareResult.comparison && compareResult.comparison.length > 0) {
      compareResult.comparison.forEach(function(row) {
        rows.push(row);
      });
    } else {
      rows.push(["אין נתונים להשוואה", "", "", "", ""]);
    }
    
    // הבטח שכל השורות באותו אורך
    var maxCols = 0;
    rows.forEach(function(r) {
      if (r.length > maxCols) maxCols = r.length;
    });
    rows = rows.map(function(r) {
      while (r.length < maxCols) {
        r.push("");
      }
      return r;
    });
    
    sheet.getRange(1, 1, rows.length, maxCols).setValues(rows);
    sheet.autoResizeColumns(1, maxCols);
    
    var url = ss.getUrl();
    Logger.log("Compare report exported to new spreadsheet: " + url);
    return url;
  } catch (e) {
    Logger.log("Error in exportCompareToSheet: " + e);
    throw e;
  }
}

/**
 * calcPastForecast - מחשבת סיכום לפגישות שהתקיימו בטווח תאריכים נתון.
 * החישוב מבוסס על השדות:
 *    pcfsystemfield559 – הכנסה מהמפגש
 *    pcfsystemfield545 – סכום לתשלום עבור הפגישה
 *    pcfsystemfield560 – רווח מהמפגש
 *
 * הפונקציה מאפשרת סינון לפי sivug:
 *    "cycle" - לפי סוג המחזור (pcfsystemfield549)
 *    "lesson" - לפי סוג ההדרכה (pcfsystemfield542, כטקסט)
 *    "branch" - לפי סניף (מומר באמצעות getBranchName מ־cycle.pcfsystemfield451)
 *
 * @param {string} startDateStr - תאריך התחלה בפורמט ISO (למשל "2025-04-01T00:00:00")
 * @param {string} endDateStr - תאריך סיום בפורמט ISO (למשל "2025-05-01T00:00:00")
 * @param {string} sivug - "cycle", "lesson" או "branch"
 * @return {object} אובייקט עם success=true, totalIncome, totalPayment, totalProfit ו-breakdown בהתאם לסיווג.
 */
function calcPastForecast(startDateStr, endDateStr, sivug) {
  Logger.log("calcPastForecast called: start=" + startDateStr + ", end=" + endDateStr + ", sivug=" + sivug);
  
  var meetings = getPastMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total past meetings: " + meetings.length);
  
  var totalIncome = 0;
  var totalPayment = 0;
  var totalProfit = 0;
  
  // לאגרגציה לפי סיווג
  var groupData = {};
  
  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    
    // **סינון לפי statuscode אחרי השליפה**:
    if (m.statuscode !== "התקיימה") {
      continue; // מדלגים על בוטלה/נדחתה וכו'
    }
    
    // לקיחת הערכים מהשדה
    var income = parseFloat(m.pcfsystemfield559 || 0);
    var payment = parseFloat(m.pcfsystemfield545 || 0);
    var profit = parseFloat(m.pcfsystemfield560 || 0);
    
    totalIncome += income;
    totalPayment += payment;
    totalProfit += profit;
    
    if (sivug === "cycle") {
      // סינון לפי סוג המחזור
      var cycleId = m.pcfsystemfield498;
      if (!cycleId) continue;
      var cycle = getCycleById(cycleId);
      if (!cycle) continue;
      var cType = parseInt(cycle.pcfsystemfield549 || "0", 10);
      var key = "";
      if (cType === 1) key = "פרטי";
      else if (cType === 2) key = "מוסדי";
      else if (cType === 3) key = "מוסדי פר ילד";
      else key = "אחר";
      
      if (!groupData[key]) {
        groupData[key] = { count: 0, income: 0, payment: 0, profit: 0 };
      }
      groupData[key].count++;
      groupData[key].income += income;
      groupData[key].payment += payment;
      groupData[key].profit += profit;
      
    } else if (sivug === "lesson") {
      var lessonTypeText = (m.pcfsystemfield542 || "").trim();
      var key = lessonTypeText || "לא מוגדר";
      
      if (!groupData[key]) {
        groupData[key] = { count: 0, income: 0, payment: 0, profit: 0 };
      }
      groupData[key].count++;
      groupData[key].income += income;
      groupData[key].payment += payment;
      groupData[key].profit += profit;
      
    } else if (sivug === "branch") {
      var cycleId = m.pcfsystemfield498;
      if (!cycleId) continue;
      var cycle = getCycleById(cycleId);
      if (!cycle) continue;
      var branchId = cycle.pcfsystemfield451;
      var branchName = getBranchName(branchId);
      var key = branchName;
      
      if (!groupData[key]) {
        groupData[key] = { count: 0, income: 0, payment: 0, profit: 0 };
      }
      groupData[key].count++;
      groupData[key].income += income;
      groupData[key].payment += payment;
      groupData[key].profit += profit;
    }
  }
  
  // המרת groupData למערך
  var breakdown = [];
  for (var k in groupData) {
    breakdown.push({
      group: k,
      count: groupData[k].count,
      income: groupData[k].income,
      payment: groupData[k].payment,
      profit: groupData[k].profit
    });
  }
  
  return {
    success: true,
    totalIncome: totalIncome,
    totalPayment: totalPayment,
    totalProfit: totalProfit,
    breakdown: breakdown
  };
}


/**
 * getPastMeetingsForRange - שליפת פגישות שהתקיימו בטווח תאריכים נתון.
 * בשונה מ-getMeetingsForRange, כאן אנו מסירים את תנאי "statuscode is-null"
 * ואפשר להוסיף תנאי "scheduledend < now()" אם נדרש.
 *
 * @param {string} startDateStr
 * @param {string} endDateStr
 * @return {array} מערך הפגישות
 */
function getPastMeetingsForRange(startDateStr, endDateStr) {
  Logger.log("getPastMeetingsForRange: " + startDateStr + " to " + endDateStr);
  var allResults = [];
  var pageNumber = 1;
  var pageSize = 200;
  var maxPages = 10;
  
  
  while (true) {
    Logger.log("Fetching past meetings page=" + pageNumber);
    var queryPayload = {
      objecttype: 6,
      page_size: pageSize,
      page_number: pageNumber,
      fields: "activityid,scheduledstart,scheduledend,pcfsystemfield485,statuscode,pcfsystemfield542,pcfsystemfield498,pcfsystemfield559,pcfsystemfield545,pcfsystemfield560",
      query:
      "(scheduledstart >= " + startDateStr + ") AND " +
      "(scheduledstart < " + endDateStr + ") AND " +
      "(pcfsystemfield498 is-not-null) AND " 
    };
    
    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) {
      Logger.log("Error or null returned for page " + pageNumber + ", stopping.");
      break;
    }
    
    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    Logger.log("Got " + chunk.length + " past meetings in this page.");
    allResults = allResults.concat(chunk);
    
    if (chunk.length < pageSize) {
      Logger.log("Less than pageSize => no more pages.");
      break;
    }
    if (pageNumber >= maxPages) {
      Logger.log("Reached maxPages=" + maxPages + ", stopping pagination.");
      break;
    }
    
    pageNumber++;
    Utilities.sleep(300);
  }
  
  Logger.log("Total past meetings after pagination: " + allResults.length);
  return allResults;
}


/**
 * calcYearlyBranchIncome - מחשבת עבור כל סניף את סך ההכנסה בין 1.9.<year> ועד 30.6.<year+1>.
 * משלבת נתוני עבר (פגישות "התקיימה") ונתוני עתיד (פגישות שלא התקיימו) ללוגיקה אחידה.
 *
 * @param {number} year - השנה שממנה מתחילים (למשל 2024). 
 *                        הטווח יהיה 1.9.year עד 30.6.(year+1).
 * @return {object} {
 *    success: true,
 *    totalIncome: סכום ההכנסות לכל הסניפים יחד,
 *    branches: [
 *      {
 *        branchName: ...,
 *        count: כמות פגישות,
 *        income: סך הכנסה,
 *        payment: סך תשלום,
 *        profit: סך רווח
 *      }, ...
 *    ]
 * }
 */
function calcYearlyBranchIncome(year) {
  // בניית תאריכי התחלה וסיום בפורמט ISO
  var startDateStr = year + "-09-01T00:00:00";          // 1 בספטמבר של אותה שנה
  var endDateStr   = (year+1) + "-06-30T23:59:59";      // 30 ביוני של השנה העוקבת

  // שליפת כל הפגישות בלי סינון סטטוס, כדי לכלול גם עבר וגם עתיד
  var meetings = getAllMeetingsForRange(startDateStr, endDateStr);
  
  // מפת צבירה לפי סניף
  var branchMap = {};
  var totalIncomeAll = 0;
  var totalPaymentAll = 0;
  var totalProfitAll = 0;

  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    
    // שליפת מזהה מחזור וסניף
    var cycleId = m.pcfsystemfield498;
    if (!cycleId) continue;
    var cycle = getCycleById(cycleId);
    if (!cycle) continue;
    var branchId = cycle.pcfsystemfield451;
    var branchName = getBranchName(branchId) || "לא צוין";

    // בדיקת statuscode
    var status = m.statuscode || "";
    // אם "התקיימה" => נשתמש בשדות שכבר מולאו (pcfsystemfield559, pcfsystemfield545, pcfsystemfield560)
    // אחרת => נשתמש בלוגיקת תחזית (getIncomePerMeeting וכד').
    var income = 0, payment = 0, profit = 0;
    if (status.trim() === "התקיימה") {
      income  = parseFloat(m.pcfsystemfield559 || 0);
      payment = parseFloat(m.pcfsystemfield545 || 0);
      profit  = parseFloat(m.pcfsystemfield560 || 0);
    } else {
      // פגישה עתידית או סטטוס אחר => לוגיקת תחזית
      // למשל: income = getIncomePerMeeting(cycle)  ...
      // או לחילופין, פונקציה אחרת אם יש.
      income  = getIncomePerMeeting(cycle);  // דוגמה
      payment = 0;                           // אם יש לוגיקת תשלום עתידי, אפשר להוסיף
      profit  = income - payment;            // או כל הגיון אחר
    }

    // צבירה
    if (!branchMap[branchName]) {
      branchMap[branchName] = {
        count: 0,
        income: 0,
        payment: 0,
        profit: 0
      };
    }
    branchMap[branchName].count++;
    branchMap[branchName].income  += income;
    branchMap[branchName].payment += payment;
    branchMap[branchName].profit  += profit;

    totalIncomeAll  += income;
    totalPaymentAll += payment;
    totalProfitAll  += profit;
  }

  // הפיכת המפה למערך
  var branchesArr = [];
  for (var bName in branchMap) {
    branchesArr.push({
      branchName: bName,
      count: branchMap[bName].count,
      income: branchMap[bName].income,
      payment: branchMap[bName].payment,
      profit: branchMap[bName].profit
    });
  }

  return {
    success: true,
    totalIncome: totalIncomeAll,
    totalPayment: totalPaymentAll,
    totalProfit: totalProfitAll,
    branches: branchesArr
  };
}


function getAllMeetingsForRange(startDateStr, endDateStr) {
  Logger.log("getAllMeetingsForRange: " + startDateStr + " to " + endDateStr);
  var allResults = [];
  var pageNumber = 1;
  var pageSize = 200;
  var maxPages = 10;

  while (true) {
    var queryPayload = {
      objecttype: 6,
      page_size: pageSize,
      page_number: pageNumber,
      fields: "activityid,scheduledstart,scheduledend,statuscode,pcfsystemfield498,pcfsystemfield559,pcfsystemfield545,pcfsystemfield560,pcfsystemfield542",
      query:
        "(scheduledstart >= " + startDateStr + ") AND " +
        "(scheduledstart <= " + endDateStr + ") AND " +
        "(pcfsystemfield498 is-not-null)"
    };

    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) {
      Logger.log("No data or error at page " + pageNumber);
      break;
    }

    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    Logger.log("Got " + chunk.length + " meetings in page " + pageNumber);
    allResults = allResults.concat(chunk);

    if (chunk.length < pageSize || pageNumber >= maxPages) {
      break;
    }
    pageNumber++;
    Utilities.sleep(300);
  }

  Logger.log("Total meetings after pagination: " + allResults.length);
  return allResults;
}



/***************************************************
 *  פונקציה לחישוב הכנסה שנתית לפי סניף
 ***************************************************/
/**
 * calculateAnnualIncomeByBranch - מחשבת את הסיכום השנתי עבור כל הפגישות (עבר/עתידי)
 * עבור סניף מסוים, בטווח שנקבע על ידי השנה הנוכחית או בהתאמה.
 * עבור כל פגישה שמוגדרת כמשויכת לסניף, נוסף גם פירוט:
 * - תאריך הפגישה (YYYY-MM-DD)
 * - שעת התחלה (HH:mm)
 * - שם המדריך (מתוך meeting.pcfsystemfield485)
 *
 * @param {string} branchId - מזהה הסניף (כפי שנשמר ב־cycle.pcfsystemfield451)
 * @return {object} אובייקט המכיל:
 *    income: סך ההכנסה,
 *    cost: סך ההוצאה,
 *    meetingCount: מספר הפגישות,
 *    profit: הרווח,
 *    meetingDetails: מערך של פרטי פגישות (date, startTime, guideName)
 */
/**
 * calcYearlyBranchIncome - מחשבת את הסיכום השנתי עבור כל הפגישות (עבר/עתידי)
 * עבור סניף מסוים, בטווח בין 1 בספטמבר לשנתיים (לפי הבחירה)
 * עבור כל פגישה, אם statuscode הוא "התקיימה", משתמשים בערכים שנרשמו בפגישה:
 *   - pcfsystemfield559 – הכנסה מהמפגש
 *   - pcfsystemfield545 – סכום לתשלום עבור הפגישה
 *   - pcfsystemfield560 – רווח מהמפגש
 * אחרת, משתמשים בלוגיקת התחזית (getIncomePerMeeting).
 * הסכום הכולל נעשה על ידי צבירת הערכים עבור כל הפגישות.
 *
 * @param {number} year - השנה שממנה מתחילים (למשל 2024).
 *                        הטווח יהיה מ־1 בספטמבר של השנה ועד 30 ביוני של השנה הבאה.
 * @return {object} אובייקט המכיל:
 *    success: true,
 *    totalIncome: סך ההכנסות,
 *    totalPayment: סך התשלומים,
 *    totalProfit: סך הרווח,
 *    meetingCount: מספר הפגישות,
 *    branches: מערך של פרטי סניפים עם count, income, payment, profit.
 */
function calculateAnnualIncomeByBranch(branchId) {
  Logger.log("calculateAnnualIncomeByBranch called for branchId: " + branchId);

  // קביעת טווח התקופה:
  // אם היום לפני ספטמבר – התקופה היא מ-1 בספטמבר של השנה שעברה ועד 30 ביוני של השנה הנוכחית.
  // אם היום בספטמבר ומעלה – התקופה היא מ-1 בספטמבר של השנה הנוכחית ועד 30 ביוני של השנה הבאה.
  var now = new Date();
  var currentYear = now.getFullYear();
  var startYear, endYear;
  if (now.getMonth() < 8) { // חודשים 0-7 (ינואר עד אוגוסט)
    startYear = currentYear - 1;
    endYear = currentYear;
  } else {
    startYear = currentYear;
    endYear = currentYear + 1;
  }
  
  var startOfPeriod = new Date(startYear, 8, 1); // 1 בספטמבר
  var endOfPeriod = new Date(endYear, 5, 30, 23, 59, 59); // 30 ביוני

  var startDateStr = toIsoString(startOfPeriod);
  var endDateStr = toIsoString(endOfPeriod);
  Logger.log("Annual period: " + startDateStr + " to " + endDateStr);

  // שליפת כל הפגישות לטווח זה – יש להשתמש בפונקציה getAllMeetingsForRange כדי לכלול את כל הפגישות
  var meetings = getAllMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total meetings for the period: " + meetings.length);

  var totalIncomeAll = 0;
  var totalPaymentAll = 0;
  var totalProfitAll = 0;
  var meetingCount = 0;
  var meetingDetails = [];
  var branchMap = {};

  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    // ודא שהפגישה משויכת למחזור
    if (!m.pcfsystemfield498) continue;
    var cycle = getCycleById(m.pcfsystemfield498);
    if (!cycle) continue;
    // בדיקה: האם המחזור שייך לסניף המבוקש
    if (cycle.pcfsystemfield451 != branchId) continue;

    meetingCount++;
    var status = m.statuscode ? m.statuscode.trim() : "";
    var income = 0, payment = 0, profit = 0;
    if (status === "התקיימה") {
      // במקרה של פגישה שהתקיימה – קריאה לערכים הקיימים בשדות
      income  = parseFloat(m.pcfsystemfield559 || 0);
      payment = parseFloat(m.pcfsystemfield545 || 0);
      profit  = parseFloat(m.pcfsystemfield560 || 0);
    } else {
      // פגישה עתידית – שימוש בלוגיקת התחזית
      income  = getIncomePerMeeting(cycle);
      payment = 0;
      profit  = income;
    }
    
    totalIncomeAll += income;
    totalPaymentAll += payment;
    totalProfitAll += profit;

    // איסוף פרטי פגישה לדיבוג/תצוגה
    var meetingDateObj = new Date(m.scheduledstart);
    var meetingDate = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var startTime = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "HH:mm");
    var guide = getGuideById(m.pcfsystemfield485);
    var guideName = guide ? guide.name : "לא נמצא";

    meetingDetails.push({
      meetingDate: meetingDate,
      startTime: startTime,
      guideName: guideName
    });

    // צבירת נתונים לפי סניף (לצורך פירוט לפי סניפים)
    var branchName = getBranchName(cycle.pcfsystemfield451) || "לא צוין";
    if (!branchMap[branchName]) {
      branchMap[branchName] = { count: 0, income: 0, payment: 0, profit: 0 };
    }
    branchMap[branchName].count++;
    branchMap[branchName].income += income;
    branchMap[branchName].payment += payment;
    branchMap[branchName].profit += profit;
  }

  // המרת branchMap למערך
  var branchesArr = [];
  for (var key in branchMap) {
    branchesArr.push({
      branchName: key,
      count: branchMap[key].count,
      income: branchMap[key].income,
      payment: branchMap[key].payment,
      profit: branchMap[key].profit
    });
  }

  var result = {
    success: true,
    income: totalIncomeAll,
    cost: totalPaymentAll,
    profit: totalProfitAll,
    meetingCount: meetingCount,
    meetingDetails: meetingDetails,
    branches: branchesArr
  };

  Logger.log("Annual Income by Branch result: " + JSON.stringify(result));
  return result;
}



function getBranchList() {
  Logger.log("getBranchList called");

  // 1. שליפת מחזורים פעילים
  var activeCycles = getActiveCycles();
  if (!activeCycles) {
    Logger.log("Error fetching active cycles");
    return null;
  }
  Logger.log("Found " + activeCycles.length + " active cycles");

  // 2. בניית רשימת סניפים ייחודית
  var uniqueBranchIds = {};
  var branches = [];

  for (var i = 0; i < activeCycles.length; i++) {
    var cycle = activeCycles[i];
    var branchId = cycle.pcfsystemfield451;

    // ודא שמזהה הסניף קיים
    if (!branchId) continue;

    // ודא שהסניף לא נמצא כבר ברשימה
    if (!uniqueBranchIds[branchId]) {
      uniqueBranchIds[branchId] = true; // הוסף את הסניף לרשימה הייחודית

      // שליפת פרטי הסניף
      var branch = getBranchById(branchId);
      if (branch) {
        branches.push({
          id: branch.customobject1009id,
          name: branch.name
        });
      }
    }
  }

  Logger.log("Found " + branches.length + " unique branches from active cycles");
  return branches;
}

function getActiveCycles() {
  Logger.log("getActiveCycles called");
  var queryPayload = {
    objecttype: 1000, // ObjectType עבור מחזורים
    page_size: 200,
    page_number: 1,
    fields: "customobject1000id,pcfsystemfield451,pcfsystemfield37,name,pcfsystemfield550,pcfsystemfield33,pcfsystemfield35", // הוספת שדות שם והכנסה
    query: "pcfsystemfield37 = 3" // שאילתה לסינון מחזורים פעילים (סטטוס 3)
  };

  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("Error fetching active cycles");
    return null;
  }

  return data.data.Data;
}

function printActiveCycles() {
  const activeCycles = getActiveCycles();
  if (!activeCycles || activeCycles.length === 0) {
    return [];
  }
  
  return activeCycles.map((cycle) => {
    return {
      Name: cycle.name || '',
      BranchName: getBranchName(cycle.pcfsystemfield451),
      Income: parseFloat(cycle.pcfsystemfield550 || 0).toFixed(2),
      CycleId: cycle.customobject1000id,
      StartDate: cycle.pcfsystemfield33 || '',
      EndDate: cycle.pcfsystemfield35 || ''
    };
  });
}

/**
 * calcBranchIncomeForChosenPeriod - מחשבת את הסיכום עבור סניף מסוים עבור טווח תאריכים נתון.
 *
 * @param {string} branchId - מזהה הסניף (כפי שנשמר ב־cycle.pcfsystemfield451)
 * @param {string} startDateStr - תאריך התחלה בפורמט ISO, לדוגמה "2025-09-01T00:00:00"
 * @param {string} endDateStr - תאריך סיום בפורמט ISO, לדוגמה "2025-06-30T23:59:59"
 * @return {object} אובייקט המכיל success, income, cost, profit, meetingCount, meetingDetails.
 */
function calcBranchIncomeForChosenPeriod(branchId, startDateStr, endDateStr) {
  Logger.log("calcBranchIncomeForChosenPeriod called for branchId: " + branchId + " from " + startDateStr + " to " + endDateStr);
  
  var meetings = getAllMeetingsForRange(startDateStr, endDateStr);
  Logger.log("Total meetings in chosen period: " + meetings.length);
  
  var totalIncome = 0, totalPayment = 0, totalProfit = 0, meetingCount = 0;
  var meetingDetails = [];
  
  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    if (!m.pcfsystemfield498) continue;
    var cycle = getCycleById(m.pcfsystemfield498);
    if (!cycle) continue;
    // רק אם המחזור שייך לסניף הנבחר
    if (cycle.pcfsystemfield451 != branchId) continue;
    
    meetingCount++;
    var status = m.statuscode ? m.statuscode.trim() : "";
    var income = 0, payment = 0, profit = 0;
    if (status === "התקיימה") {
      income  = parseFloat(m.pcfsystemfield559 || 0);
      payment = parseFloat(m.pcfsystemfield545 || 0);
      profit  = parseFloat(m.pcfsystemfield560 || 0);
    } else {
      income  = getIncomePerMeeting(cycle);
      payment = 0;
      profit  = income;
    }
    
    totalIncome += income;
    totalPayment += payment;
    totalProfit += profit;
    
    var meetingDateObj = new Date(m.scheduledstart);
    var meetingDate = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var startTime = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "HH:mm");
    
    var guide = getGuideById(m.pcfsystemfield485);
    var guideName = guide ? guide.name : "לא נמצא";
    
    meetingDetails.push({
      meetingDate: meetingDate,
      startTime: startTime,
      guideName: guideName
    });
  }
  
  Logger.log("Total income: " + totalIncome + ", total payment: " + totalPayment + ", total profit: " + totalProfit + ", meetings: " + meetingCount);
  
  return {
    success: true,
    income: totalIncome,
    cost: totalPayment,
    profit: totalProfit,
    meetingCount: meetingCount,
    meetingDetails: meetingDetails
  };
}

/**
 * getMeetingsForCycle - שליפת פגישות עבור מחזור מסוים בטווח תאריכים נתון.
 * @param {string} cycleId - מזהה המחזור (למשל, הערך ב־customobject1000id)
 * @param {string} startDateStr - תאריך התחלה בפורמט ISO
 * @param {string} endDateStr - תאריך סיום בפורמט ISO
 * @return {Array} מערך הפגישות עבור המחזור.
 */
function getMeetingsForCycle(cycleId, startDateStr, endDateStr) {
  Logger.log("getMeetingsForCycle called for cycleId: " + cycleId + " from " + startDateStr + " to " + endDateStr);
  var allResults = [];
  var pageNumber = 1;
  var pageSize = 200;
  var maxPages = 10;
  while (true) {
    var queryPayload = {
      objecttype: 6,
      page_size: pageSize,
      page_number: pageNumber,
      fields: "activityid, scheduledstart, scheduledend, pcfsystemfield485, pcfsystemfield542, pcfsystemfield498, statuscode, pcfsystemfield559, pcfsystemfield545, pcfsystemfield560",
      query:
        "(scheduledstart >= " + startDateStr + ") AND " +
        "(scheduledstart < " + endDateStr + ") AND " +
        "(pcfsystemfield498 = '" + cycleId + "')"
    };
    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) {
      Logger.log("Error or null returned for cycle " + cycleId + ", page " + pageNumber);
      break;
    }
    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    Logger.log("Got " + chunk.length + " meetings for cycle " + cycleId + " in page " + pageNumber);
    allResults = allResults.concat(chunk);
    if (chunk.length < pageSize || pageNumber >= maxPages) break;
    pageNumber++;
    Utilities.sleep(300);
  }
  Logger.log("Total meetings for cycle " + cycleId + ": " + allResults.length);
  return allResults;
}

/**
 * getCyclesForBranch - שליפת מחזורים (ObjectType=1000) ששייכים לסניף מסוים.
 * @param {string} branchId - מזהה הסניף (כפי שמופיע ב־cycle.pcfsystemfield451)
 * @return {Array} מערך המחזורים השייכים לסניף.
 */
function getCyclesForBranch(branchId) {
  Logger.log("getCyclesForBranch called for branchId: " + branchId);
  var queryPayload = {
    objecttype: 1000,
    page_size: 200,
    page_number: 1,
    fields: "customobject1000id,pcfsystemfield451",
    query: "(pcfsystemfield451 = '" + branchId + "')"
  };
  var cycles = [];
  var pageNumber = 1;
  var pageSize = 200;
  var maxPages = 10;
  while (true) {
    queryPayload.page_number = pageNumber;
    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!data) break;
    var chunk = (data.data && data.data.Data) ? data.data.Data : [];
    cycles = cycles.concat(chunk);
    if (chunk.length < pageSize || pageNumber >= maxPages) break;
    pageNumber++;
    Utilities.sleep(300);
  }
  Logger.log("Found " + cycles.length + " cycles for branch " + branchId);
  return cycles;
}

/**
 * getMeetingsForBranchRange - שליפת כל הפגישות בטווח תאריכים עבור סניף מסוים.
 * היא מושכת את כל המחזורים השייכים לסניף ואז עבור כל מחזור שולפת את הפגישות שלו.
 *
 * @param {string} branchId - מזהה הסניף.
 * @param {string} startDateStr - תאריך התחלה בפורמט ISO.
 * @param {string} endDateStr - תאריך סיום בפורמט ISO.
 * @return {Array} מערך הפגישות עבור הסניף.
 */
function getMeetingsForBranchRange(branchId, startDateStr, endDateStr) {
  Logger.log("getMeetingsForBranchRange called for branchId: " + branchId + " from " + startDateStr + " to " + endDateStr);
  var cycles = getCyclesForBranch(branchId);
  var allMeetings = [];
  for (var i = 0; i < cycles.length; i++) {
    var cycle = cycles[i];
    var cycleId = cycle.customobject1000id;  // או cycle.id בהתאם למבנה
    if (!cycleId) continue;
    var meetings = getMeetingsForCycle(cycleId, startDateStr, endDateStr);
    allMeetings = allMeetings.concat(meetings);
    Utilities.sleep(300);
  }
  Logger.log("Total meetings for branch " + branchId + ": " + allMeetings.length);
  return allMeetings;
}


function calcBranchIncomeForPeriod(branchId, startDateStr, endDateStr) {
  Logger.log("calcBranchIncomeForPeriod called for branch " + branchId + " from " + startDateStr + " to " + endDateStr);
  
  // שליפת כל הפגישות של הסניף בטווח הנתון – באמצעות getMeetingsForBranchRange
  var meetings = getMeetingsForBranchRange(branchId, startDateStr, endDateStr);
  Logger.log("Total meetings for branch in chosen period: " + meetings.length);
  
  var totalIncome = 0;
  var totalPayment = 0;
  var totalProfit = 0;
  var meetingCount = 0;
  var meetingDetails = [];
  
  for (var i = 0; i < meetings.length; i++) {
    var m = meetings[i];
    
    // ודא שהפגישה משויכת למחזור
    if (!m.pcfsystemfield498) continue;
    var cycle = getCycleById(m.pcfsystemfield498);
    if (!cycle) continue;
    
    // סינון: אם קיימת סטטוס (לא ריק) – נכלול רק אם הוא "התקיימה" או "4"
    var status = m.statuscode ? m.statuscode.trim() : "";
    if (status !== "" && status !== "התקיימה" && status !== "4") {
      continue;
    }
    
    meetingCount++;
    var income = 0, payment = 0, profit = 0;
    
    if (status === "התקיימה" || status === "4") {
      // עבור פגישות שהתקיימו – משתמשים בערכים שנרשמו
      income  = parseFloat(m.pcfsystemfield559 || 0);
      payment = parseFloat(m.pcfsystemfield545 || 0);
      profit  = parseFloat(m.pcfsystemfield560 || 0);
    } else {
      // עבור פגישות עתידיות – מחשבים לפי לוגיקת התחזית
      income  = getIncomePerMeeting(cycle);
      payment = calculateMeetingCost(m);
      profit  = income - payment;
    }
    
    totalIncome += income;
    totalPayment += payment;
    totalProfit += profit;
    
    // איסוף פרטי הפגישה להצגה – כולל תאריך, שעת התחלה, שם מדריך והערכים המחושבים
    var meetingDateObj = new Date(m.scheduledstart);
    var meetingDate = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var startTime = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "HH:mm");
    var guide = getGuideById(m.pcfsystemfield485);
    var guideName = guide ? guide.name : "לא נמצא";
    
    meetingDetails.push({
      meetingDate: meetingDate,
      startTime: startTime,
      guideName: guideName,
      income: income,
      cost: payment,
      profit: profit
    });
  }
  
  // מיין את פרטי הפגישות לפי תאריך בסדר עולה
  meetingDetails.sort(function(a, b) {
    return new Date(a.meetingDate) - new Date(b.meetingDate);
  });
  
  Logger.log("Total income: " + totalIncome + ", Total payment: " + totalPayment + ", Total profit: " + totalProfit + ", Meetings count: " + meetingCount);
  
  return {
    success: true,
    income: totalIncome,
    cost: totalPayment,
    profit: totalProfit,
    meetingCount: meetingCount,
    meetingDetails: meetingDetails
  };
}




function testBranchIncomeForChosenPeriod() {
  var branchId = "b8cc6be7-1bc7-4d45-ad6b-dac8aec85bb9";
  var startDateStr = "2025-03-18T00:00:00";
  var endDateStr = "2025-03-18T23:59:59";
  
  var result = calcBranchIncomeForPeriod(branchId, startDateStr, endDateStr);
  Logger.log("Annual Branch Income for branch " + branchId + " from " + startDateStr + " to " + endDateStr + ":\n" + JSON.stringify(result));
}


/**
 * testPastMeetingsDebug - פונקציית עזר לדיבוג חישוב פגישות עבר
 * מדמה הרצה של calcPastForecast עם פרמטרים לדוגמה,
 * ומדפיסה את התוצאה ללוג של Apps Script.
 */
function testPastMeetingsDebug() {
  // הגדרת טווח תאריכים בפורמט ISO, לדוגמה לחודש מרץ 2025
  var startDateStr = "2025-03-01T00:00:00";
  var endDateStr   = "2025-03-31T00:00:00";
  
  // בחירת סיווג: "cycle", "lesson" או "branch"
  var sivug = "branch";
  
  // קריאה לחישוב
  var result = calcPastForecast(startDateStr, endDateStr, sivug);
  
  // הדפסת התוצאה ללוג
  Logger.log("Past Forecast from " + startDateStr + " to " + endDateStr +
             " (sivug=" + sivug + "): " + JSON.stringify(result));
}

function testCompareMonthlyForecast_March_April_2025() {
  var month1 = "2025-02";  // March 2025
  var month2 = "2025-03";  // April 2025
  
  // Call the function that compares monthly forecasts by branch.
  var result = compareMonthlyForecastByBranch(month1, month2);
  
  // Log the result for debugging.
  Logger.log("Comparison between " + month1 + " and " + month2 + ": " + JSON.stringify(result));
}


/**
 * analyzeComparisonByBranch - Analyzes the branch comparison data using OpenAI.
 * It builds a prompt from the compareResult data, sends a request to the OpenAI API,
 * and returns an in‑depth analysis.
 *
 * @param {object} compareResult - The comparison result object from compareMonthlyForecastByBranch.
 * @return {object} An object with a success flag and an analysis string.
 */
function analyzeComparisonByBranch(compareResult) {
  // Build a prompt based on the compare result data
  
  var prompt = "Analyze the following branch comparison data and provide an in‑depth explanation for the observed differences between the months:\n\n" 
               + JSON.stringify(compareResult, null, 2);
  
   
  var analysis = callOpenAIApi(prompt);
  return { success: true, analysis: analysis };
  
}




/**
 * callOpenAIApi - Calls the OpenAI API with the given prompt and returns the result text.
 *
 * @param {string} prompt - The prompt text to send to OpenAI.
 * @return {string} The analysis text returned by the API.
 */
function callOpenAIApi(prompt) {
  var apiUrl = "https://api.openai.com/v1/chat/completions";
  

  // Prepare the payload using the Chat Completion format.
  var payload = {
  "model": "gpt-4-turbo",
  "messages": [
    {
      "role": "user",
      "content": prompt
    }
  ],
  "max_tokens": 750,
  "temperature": 0.3,
  "top_p": 1.0,
  "frequency_penalty": 0.2,
  "presence_penalty": 0.1
};


  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  
  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var json = JSON.parse(response.getContentText());
    Logger.log("OpenAI API response: " + JSON.stringify(json));
    
    // Check if an error was returned by the API
    if (json.error) {
      return "OpenAI API error: " + json.error.message;
    }
    
    // Make sure the expected properties exist before calling trim()
    if (json && json.choices && json.choices.length > 0 && 
        json.choices[0].message && json.choices[0].message.content) {
      return json.choices[0].message.content.trim();
    } else {
      return "No analysis returned from OpenAI API.";
    }
  } catch (error) {
    Logger.log("Error calling OpenAI API: " + error);
    return "Error calling OpenAI API: " + error.toString();
  }
}

function testAnnualForecastHerzog() {
  var branchId = "d9b621ba-e10d-43a5-9edd-6a33b4b5709a";
  var result = calculateAnnualIncomeByBranch(branchId);
  Logger.log("Annual forecast for גימנסיה הרצוג (branchId: " + branchId + "): " + JSON.stringify(result));
}





/**
 * פונקציה להרצת תחזית שבועית עבור אפריל 2025, שבוע 1,
 * לפי סיווג "cycle" (לפי סוג המחזור).
 * הפונקציה מדפיסה ללוג את התוצאה המוחזרת.
 */
function testWeeklyForecastApril2025Week1() {
  var monthStr = "2025-04";
  var weekStr = "1";
  var sivug = "cycle";  // ניתן לשנות ל-"lesson" או "branch" לפי הצורך

  var result = calcMonthlyForecastSS(monthStr, sivug);
  Logger.log("Weekly Forecast for April 2025, Week 1 (" + sivug + "): " + JSON.stringify(result));
}

function getEndedCycles(startDate, endDate) {
  Logger.log("getEndedCycles called with date range: " + startDate + " to " + endDate);
  var queryPayload = {
    objecttype: 1000,
    page_size: 200,
    page_number: 1,
    fields: "customobject1000id,pcfsystemfield451,pcfsystemfield37,name,pcfsystemfield550,pcfsystemfield35,pcfsystemfield33",
    query: "(pcfsystemfield37 = 4) AND " +
          "(pcfsystemfield35 >= " + startDate + "T00:00:00) AND " +
          "(pcfsystemfield35 <= " + endDate + "T23:59:59)"
  };

  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("Error fetching ended cycles");
    return null;
  }

  return data.data.Data;
}

function printEndedCycles(startDate, endDate) {
  const endedCycles = getEndedCycles(startDate, endDate);
  if (!endedCycles || endedCycles.length === 0) {
    return [];
  }
  
  return endedCycles.map((cycle) => {
    return {
      Name: cycle.name || '',
      CycleId: cycle.customobject1000id,
      StartDate: cycle.pcfsystemfield33,
      EndDate: cycle.pcfsystemfield35
    };
  });
}

function testPrintActiveCycles() {
  const cycles = printActiveCycles();
  Logger.log("Active Cycles:", JSON.stringify(cycles, null, 2));
}

/**
 * getInstitutionalOrders - שליפת הזמנות מוסדיות בסטטוס מעקב וגבייה והסתיימה
 * @return {Array} מערך של הזמנות מוסדיות עם פרטיהן
 */
function getInstitutionalOrders() {
  Logger.log("getInstitutionalOrders called");
  var queryPayload = {
    objecttype: 1005,  // ObjectType עבור הזמנות מוסדיות
    page_size: 200,
    page_number: 1,
    fields: "customobject1005id,name,pcfsystemfield192,pcfsystemfield182,pcfsystemfield181",  // שדות: מזהה, שם, תשלום כולל, סטטוס, סניף
    query: "(pcfsystemfield182 = 10) OR (pcfsystemfield182 = 13)"  // סינון לפי סטטוס מעקב וגבייה (10) או הסתיים (13)
  };

  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("Error fetching institutional orders");
    return null;
  }

  return data.data.Data;
}

/**
 * printInstitutionalOrders - מחזירה את ההזמנות המוסדיות בפורמט מתאים לתצוגה
 * @return {Array} מערך של הזמנות מוסדיות מעובדות לתצוגה
 */
function printInstitutionalOrders() {
  const orders = getInstitutionalOrders();
  if (!orders || orders.length === 0) {
    return [];
  }
  
  return orders.map((order) => {
    const payment = parseFloat(order.pcfsystemfield192 || 0);
    const statusNum = order.pcfsystemfield182;
    let statusText = "";
    
    // המרה ממספר לטקסט
    if (statusNum === "מעקב וגבייה") {
      statusText = "מעקב וגבייה";
    } else if (statusNum === "הסתיים") {
      statusText = "הסתיים";
    } else {
      statusText = statusNum; // במקרה שיש ערך לא צפוי
    }

    return {
      Name: order.name || '',
      OrderId: order.customobject1005id,
      TotalPayment: payment.toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      BranchName: getBranchName(order.pcfsystemfield181),
      Status: statusText,
      RawStatus: statusNum // שומר את הסטטוס המספרי המקורי
    };
  });
}

/**
 * testInstitutionalOrders - פונקציית בדיקה להזמנות מוסדיות
 * מדפיסה את כל הפרטים כולל סכומים לפי סטטוס
 */
function testInstitutionalOrders() {
  const orders = getInstitutionalOrders();
  Logger.log("Raw orders from API:", orders);

  const processedOrders = printInstitutionalOrders();
  Logger.log("Processed orders:", processedOrders);

  // חישוב סכומים לפי סטטוס
  let totalTracking = 0;
  let totalCompleted = 0;
  
  processedOrders.forEach(order => {
    const amount = parseFloat(order.TotalPayment.replace(/[₪,\s]/g, ''));
    Logger.log("Processing order:", {
      name: order.Name,
      status: order.Status,
      rawStatus: order.RawStatus,
      amount: amount
    });

    if (order.RawStatus === "10") {
      totalTracking += amount;
      Logger.log("Added to tracking total:", amount, "New total:", totalTracking);
    } else if (order.RawStatus === "13") {
      totalCompleted += amount;
      Logger.log("Added to completed total:", amount, "New total:", totalCompleted);
    }
  });

  Logger.log("Final Totals:", {
    tracking: totalTracking,
    completed: totalCompleted,
    total: totalTracking + totalCompleted
  });
}

/**
 * compareActiveCyclesEndDates - משווה בין תאריך סיום המחזור לתאריך הפגישה האחרונה
 * @return {Array} מערך של מחזורים שבהם יש פער בין התאריכים
 */
function compareActiveCyclesEndDates() {
  Logger.log("compareActiveCyclesEndDates called");
  
  var cyclesPayload = {
    objecttype: 1000,
    page_size: 200,
    page_number: 1,
    fields: "customobject1000id,name,pcfsystemfield35,pcfsystemfield37",
    query: "pcfsystemfield37 = 3"
  };

  var cyclesData = sendRequestWithRetry(QUERY_API_URL, cyclesPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!cyclesData || !cyclesData.data || !cyclesData.data.Data) {
    Logger.log("Error fetching active cycles");
    return null;
  }

  var activeCycles = cyclesData.data.Data;
  var results = [];

  for (var i = 0; i < activeCycles.length; i++) {
    var cycle = activeCycles[i];
    var cycleId = cycle.customobject1000id;
    var cycleEndDate = cycle.pcfsystemfield35;

    if (!cycleEndDate) {
      Logger.log("Cycle " + cycleId + " has no end date");
      continue;
    }

    var meetingsPayload = {
      objecttype: 6,
      page_size: 200,
      page_number: 1,
      fields: "scheduledstart,pcfsystemfield498",
      query: "pcfsystemfield498 = '" + cycleId + "'",
      sort: [{ field: "scheduledstart", direction: "desc" }]
    };

    var meetingsData = sendRequestWithRetry(QUERY_API_URL, meetingsPayload, MAX_RETRIES, RETRY_DELAY_MS);
    if (!meetingsData || !meetingsData.data || !meetingsData.data.Data || meetingsData.data.Data.length === 0) {
      Logger.log("No meetings found for cycle: " + cycleId);
      continue;
    }

    var meetings = meetingsData.data.Data;
    meetings.sort(function(a, b) {
      return new Date(b.scheduledstart) - new Date(a.scheduledstart);
    });

    var lastMeeting = meetings[0];
    var lastMeetingDate = lastMeeting.scheduledstart;

    Logger.log("Cycle: " + cycle.name);
    Logger.log("End Date: " + cycleEndDate);
    Logger.log("Last Meeting Date: " + lastMeetingDate);

    var cycleEndDateTime = new Date(cycleEndDate);
    var lastMeetingDateTime = new Date(lastMeetingDate);

    if (Math.abs(cycleEndDateTime - lastMeetingDateTime) > 24 * 60 * 60 * 1000) {
      results.push({
        CycleName: cycle.name,
        CycleId: cycleId,
        EndDate: cycleEndDate,
        LastMeetingDate: lastMeetingDate,
        DaysDifference: Math.round(Math.abs(cycleEndDateTime - lastMeetingDateTime) / (24 * 60 * 60 * 1000))
      });
    }

    Utilities.sleep(100);
  }

  return results;
}

/**
 * printCyclesEndDateComparison - מדפיסה את תוצאות ההשוואה בפורמט מסודר
 * @return {Array} מערך מעובד של תוצאות ההשוואה
 */
function printCyclesEndDateComparison() {
  const results = compareActiveCyclesEndDates();
  if (!results || results.length === 0) {
    return [];
  }
  
  return results.map(result => {
    const endDate = new Date(result.EndDate);
    const lastMeetingDate = new Date(result.LastMeetingDate);
    
    return {
      Name: result.CycleName,
      CycleId: result.CycleId,
      EndDate: endDate.toLocaleDateString('he-IL'),
      LastMeeting: lastMeetingDate.toLocaleDateString('he-IL'),
      DaysDiff: result.DaysDifference,
      RawLastMeetingDate: result.LastMeetingDate // שומר את התאריך המקורי לצורך העדכון
    };
  });
}

/**
 * updateCycleEndDate - מעדכן את תאריך הסיום של מחזור
 * @param {string} cycleId - מזהה המחזור
 * @param {string} newEndDate - תאריך הסיום החדש
 * @return {object} תוצאת העדכון
 */
function updateCycleEndDate(cycleId, newEndDate) {
  Logger.log("updateCycleEndDate called for cycle " + cycleId + " with new date " + newEndDate);
  
  var url = API_URL + "/1000/" + cycleId;
  var payload = {
    "pcfsystemfield35": newEndDate
  };
  
  var options = {
    method: 'PUT',
    contentType: 'application/json',
    headers: {
      tokenid: FIREBERRY_API_KEY,
      accept: 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      // ניקוי המטמון של המחזור כדי שנקבל את הנתונים המעודכנים בפעם הבאה
      if (cycleCache[cycleId]) {
        delete cycleCache[cycleId];
      }
      return { success: true, message: "תאריך הסיום עודכן בהצלחה" };
    } else {
      return { success: false, message: "שגיאה בעדכון תאריך הסיום: " + response.getContentText() };
    }
  } catch (error) {
    Logger.log("Error updating cycle end date: " + error);
    return { success: false, message: "שגיאה בעדכון תאריך הסיום: " + error.toString() };
  }
}

/**
 * updateAllCyclesEndDates - מעדכן את כל תאריכי הסיום של המחזורים לפי תאריך הפגישה האחרונה
 * @return {object} תוצאת העדכון הכולל
 */
function updateAllCyclesEndDates() {
  Logger.log("updateAllCyclesEndDates called");
  
  var results = compareActiveCyclesEndDates();
  if (!results || results.length === 0) {
    return { success: true, message: "לא נמצאו מחזורים לעדכון", updatedCount: 0 };
  }

  var successCount = 0;
  var errorCount = 0;
  var errors = [];

  for (var i = 0; i < results.length; i++) {
    var cycle = results[i];
    try {
      var updateResult = updateCycleEndDate(cycle.CycleId, cycle.LastMeetingDate);
      if (updateResult.success) {
        successCount++;
      } else {
        errorCount++;
        errors.push("שגיאה בעדכון מחזור " + cycle.CycleName + ": " + updateResult.message);
      }
      // הוספת השהייה קטנה בין העדכונים כדי לא להעמיס על ה-API
      Utilities.sleep(300);
    } catch (error) {
      errorCount++;
      errors.push("שגיאה בעדכון מחזור " + cycle.CycleName + ": " + error.toString());
    }
  }

  return {
    success: true,
    message: "הושלם עדכון תאריכי סיום. עודכנו בהצלחה: " + successCount + " מחזורים. נכשלו: " + errorCount + " מחזורים.",
    updatedCount: successCount,
    errorCount: errorCount,
    errors: errors
  };
}

/**
 * getMeetingDetails - מביא את כל הפרטים הרלוונטיים על פגישה
 * @param {string} meetingId - מזהה הפגישה
 * @return {object} אובייקט עם כל פרטי הפגישה
 */
function getMeetingDetails(meetingId) {
  Logger.log("getMeetingDetails called for meeting: " + meetingId);
  
  var url = API_URL + "/6/" + meetingId;
  var data = sendGetWithRetry(url, MAX_RETRIES, RETRY_DELAY_MS);
  
  if (!data || !data.data || !data.data.Record) {
    Logger.log("Failed to fetch meeting details");
    return null;
  }
  
  var meeting = data.data.Record;
  var cycleId = meeting.pcfsystemfield498;
  var cycle = cycleId ? getCycleById(cycleId) : null;
  var guideId = meeting.pcfsystemfield485;
  var guide = guideId ? getGuideById(guideId) : null;
  
  // חישוב תחזית אם אין נתונים בפועל
  var income = 0, cost = 0, profit = 0;
  
  // בדיקת סטטוס - תומך גם בערך מספרי וגם בטקסט
  var status = meeting.statuscode;
  var isCompleted = false;
  
  if (status) {
    // המרה לסטרינג רק אם זה לא סטרינג
    var statusStr = typeof status === 'string' ? status : status.toString();
    isCompleted = statusStr === "3" || statusStr === "התקיימה";
  }
  
  if (isCompleted) {
    income = parseFloat(meeting.pcfsystemfield559 || 0);
    cost = parseFloat(meeting.pcfsystemfield545 || 0);
    profit = parseFloat(meeting.pcfsystemfield560 || 0);
  } else if (cycle) {
    income = getIncomePerMeeting(cycle);
    cost = calculateMeetingCost(meeting);
    profit = income - cost;
  }

  // המרת סטטוס מספרי לטקסט
  var statusText = "";
  if (status) {
    var statusStr = typeof status === 'string' ? status : status.toString();
    switch(statusStr) {
      case "3":
      case "התקיימה":
        statusText = "התקיימה";
        break;
      case "1":
      case "בוטלה":
        statusText = "בוטלה";
        break;
      case "4":
      case "נדחתה":
        statusText = "נדחתה";
        break;
      case "2":
      case "לא בשימוש":
        statusText = "לא בשימוש";
        break;
      default:
        statusText = statusStr;
    }
  } else {
    statusText = "פעיל";
  }

  return {
    id: meetingId,
    name: meeting.name || "",
    startDate: meeting.scheduledstart || "",
    endDate: meeting.scheduledend || "",
    status: statusText,
    rawStatus: status, // שמירת הערך המקורי
    lessonType: meeting.pcfsystemfield542 || "",
    guideName: guide ? guide.name : "לא ידוע",
    cycleName: cycle ? cycle.name : "לא ידוע",
    income: income,
    cost: cost,
    profit: profit,
    branchName: cycle ? getBranchName(cycle.pcfsystemfield451) : "לא ידוע"
  };
}

/**
 * searchMeetings - מחפש פגישות לפי פרמטרים שונים
 * @param {object} params - פרמטרים לחיפוש (תאריכים, סטטוס, מדריך, מחזור וכו')
 * @return {Array} מערך של פגישות שנמצאו
 */
function searchMeetings(params) {
  Logger.log("searchMeetings called with params: " + JSON.stringify(params));
  
  var query = [];
  
  if (params.startDate) {
    Logger.log("Adding startDate filter: " + params.startDate);
    query.push("(scheduledstart >= " + params.startDate + ")");
  }
  if (params.endDate) {
    Logger.log("Adding endDate filter: " + params.endDate);
    query.push("(scheduledstart < " + params.endDate + ")");
  }

  // עדכון הטיפול בסטטוס - תמיכה גם בערכים מספריים וגם בטקסט
  if (params.status) {
    Logger.log("Adding status filter: " + params.status);
    switch(params.status) {
      case "פעיל":
        // פגישות עתידיות - אין להן סטטוס
        query.push("(statuscode is-null)");
        break;
      case "התקיימה":
        query.push("((statuscode = 'התקיימה') OR (statuscode = '3'))");
        break;
      case "בוטלה":
        query.push("((statuscode = 'בוטלה') OR (statuscode = '1'))");
        break;
      case "נדחתה":
        query.push("((statuscode = 'נדחתה') OR (statuscode = '4'))");
        break;
      case "לא בשימוש":
        query.push("((statuscode = 'לא בשימוש') OR (statuscode = '2'))");
        break;
    }
  }

  if (params.guideId) {
    Logger.log("Adding guideId filter: " + params.guideId);
    query.push("(pcfsystemfield485 = '" + params.guideId + "')");
  }
  if (params.cycleId) {
    Logger.log("Adding cycleId filter: " + params.cycleId);
    query.push("(pcfsystemfield498 = '" + params.cycleId + "')");
  }
  
  // הוספת תנאי חובה שקיים מחזור
  query.push("(pcfsystemfield498 is-not-null)");
  
  var queryPayload = {
    objecttype: 6,
    page_size: 50,
    page_number: 1,
    fields: "activityid,scheduledstart,scheduledend,pcfsystemfield485,statuscode,pcfsystemfield542,pcfsystemfield498,pcfsystemfield559,pcfsystemfield545,pcfsystemfield560",
    query: query.join(" AND ")
  };
  
  Logger.log("Final query payload: " + JSON.stringify(queryPayload));
  
  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("No data returned from API");
    return [];
  }
  
  Logger.log("Found " + data.data.Data.length + " meetings");
  return data.data.Data.map(meeting => {
    var details = getMeetingDetails(meeting.activityid);
    Logger.log("Processed meeting: " + meeting.activityid);
    return details;
  });
}

/**
 * processChatQuery - מעבד שאילתת צ'אט ומחזיר תשובה
 * @param {string} query - השאילתה מהמשתמש
 * @return {object} תשובת הצ'אטבוט
 */
function processChatQuery(query) {
  Logger.log("processChatQuery called with query: " + query);
  
  try {
    // First prompt to analyze the query and determine the time range
    var analysisPrompt = `You are an AI assistant for a meeting management system that uses Fireberry API.
Your task is to analyze the user's query and determine the appropriate time range for searching meetings.

User query in Hebrew: "${query}"

Current date and time for reference: ${new Date().toISOString()}

Respond with a valid JSON object in this exact format:
{
  "queryType": "time_range" | "specific_meetings" | "need_clarification",
  "timeRange": {
    "startDate": "YYYY-MM-DDT00:00:00",
    "endDate": "YYYY-MM-DDT23:59:59"
  },
  "filters": {
    "status": null | "התקיימה" | "בוטלה" | "נדחתה" | "לא בשימוש" | "פעיל",
    "guideId": null | "string",
    "cycleId": null | "string"
  },
  "clarificationNeeded": false | true,
  "clarificationQuestion": "string" | null,
  "explanation": "string",
  "useMemory": boolean
}`;

    var analysisResponse = callOpenAIApi(analysisPrompt);
    var analysis = JSON.parse(cleanJsonResponse(analysisResponse));
    
    if (analysis.clarificationNeeded) {
      return {
        success: true,
        answer: analysis.clarificationQuestion,
        needsClarification: true
      };
    }

    var searchParams = {
      startDate: analysis.timeRange.startDate,
      endDate: analysis.timeRange.endDate,
      ...analysis.filters
    };
    
    // חיפוש בזיכרון אם רלוונטי
    let memorizedMeetings = [];
    if (analysis.useMemory) {
      memorizedMeetings = searchMemory('meetings', meeting => {
        const meetingDate = new Date(meeting.startDate);
        const startDate = new Date(searchParams.startDate);
        const endDate = new Date(searchParams.endDate);
        return meetingDate >= startDate && meetingDate <= endDate &&
               (!searchParams.status || meeting.status === searchParams.status);
      });
    }
    
    // חיפוש בזמן אמת
    var currentMeetings = searchMeetings(searchParams);
    
    // שמירת התוצאות החדשות בזיכרון
    currentMeetings.forEach(meeting => {
      saveToMemory(meeting, 'meetings', `${meeting.id}_${meeting.startDate}`);
    });
    
    // שילוב התוצאות
    var allMeetings = [...currentMeetings];
    if (memorizedMeetings.length > 0) {
      // הוספת פגישות מהזיכרון שלא נמצאו בחיפוש הנוכחי
      memorizedMeetings.forEach(memMeeting => {
        if (!allMeetings.some(m => m.id === memMeeting.id)) {
          allMeetings.push(memMeeting);
        }
      });
    }
    
    Logger.log(`Found ${currentMeetings.length} current meetings and ${memorizedMeetings.length} from memory`);
    
    // יצירת תשובה מותאמת
    var responsePrompt = `You are a helpful assistant for a meeting management system.
Analyze these meetings and provide a VERY CONCISE response in Hebrew.

User query: "${query}"
Analysis explanation: ${analysis.explanation}
Current meetings found: ${currentMeetings.length}
Memorized meetings found: ${memorizedMeetings.length}
Time range analyzed: ${analysis.timeRange.startDate} to ${analysis.timeRange.endDate}
Meetings data: ${JSON.stringify(allMeetings)}

Response Guidelines:
1. Write ONLY in Hebrew
2. Keep responses extremely short and direct
3. For questions about meeting counts, respond with just the number and basic context
4. For questions about specific meetings, provide only the directly requested information
5. Do not list individual meetings unless specifically asked
6. Do not provide breakdowns or summaries unless specifically asked
7. Focus only on answering exactly what was asked

Make the response as brief as possible while still being clear and accurate.`;

    var formattedResponse = callOpenAIApi(responsePrompt);
    
    return {
      success: true,
      answer: formattedResponse.trim(),
      data: allMeetings,
      analysis: analysis
    };
    
  } catch (error) {
    Logger.log("Error in processChatQuery: " + error + "\nStack: " + error.stack);
    return {
      success: false,
      answer: "מצטער, נתקלתי בבעיה בעיבוד השאילתה: " + error.message,
      data: null
    };
  }
}

/**
 * cleanJsonResponse - מנקה את התשובה מ-OpenAI ומוודא שהיא JSON תקין
 * @param {string} response - התשובה המקורית מ-OpenAI
 * @return {string} תשובה מנוקה
 */
function cleanJsonResponse(response) {
  if (!response) return "{}";
  
  // הסר תגי קוד אם קיימים
  response = response.replace(/```json\s*/g, "")
                    .replace(/```\s*/g, "")
                    .trim();
  
  // הסר שורות חדשות מיותרות
  response = response.replace(/\n\s*\n/g, "\n")
                    .replace(/^\s+|\s+$/g, "");
  
  // נקה תווים מיוחדים
  response = response.replace(/[\u200B-\u200D\uFEFF]/g, "");
  
  // וודא שיש סוגריים מסולסלים בהתחלה ובסוף
  if (!response.startsWith("{")) {
    response = "{" + response;
  }
  if (!response.endsWith("}")) {
    response = response + "}";
  }
  
  // נסה לפרסר ולהחזיר JSON מפורמט
  try {
    return JSON.stringify(JSON.parse(response));
  } catch (e) {
    Logger.log("Error parsing cleaned response: " + e);
    Logger.log("Cleaned response was: " + response);
    throw new Error("Could not parse OpenAI response as JSON");
  }
}

/**
 * testChatbotQuery - פונקציית בדיקה מקיפה לצ'אטבוט
 * בודקת מספר תרחישים שונים ומדפיסה את התוצאות
 */
function testChatbotQuery() {
  Logger.log("=== Starting Chatbot Test ===");
  
  // מערך של שאילתות לבדיקה
  var testQueries = [
    "כמה פגישות מתוכננות להיום?",
    "מי המדריך בפגישה הבאה?",
    "הראה לי את כל הפגישות שהתקיימו בחודש האחרון"
  ];
  
  for (var i = 0; i < testQueries.length; i++) {
    var query = testQueries[i];
    Logger.log("\n=== Testing Query: " + query + " ===");
    
    try {
      // קריאה לפונקציית העיבוד
      var result = processChatQuery(query);
      
      // הדפסת התוצאות המפורטות
      Logger.log("Success: " + result.success);
      Logger.log("Answer from Bot: " + result.answer);
      
      if (result.data) {
        if (Array.isArray(result.data)) {
          Logger.log("Total Meetings Found: " + result.data.length);
          
          // סיכום לפי סטטוס
          var statusCount = {};
          result.data.forEach(function(meeting) {
            var status = meeting.status || "לא מוגדר";
            statusCount[status] = (statusCount[status] || 0) + 1;
          });
          
          Logger.log("Status Breakdown:");
          Object.keys(statusCount).forEach(function(status) {
            Logger.log("- " + status + ": " + statusCount[status]);
          });
          
          // הדפסת פרטי הפגישה הראשונה לדוגמה
          if (result.data.length > 0) {
            Logger.log("\nExample Meeting Details:");
            Logger.log(JSON.stringify(result.data[0], null, 2));
          }
        } else {
          Logger.log("Single Meeting Details:");
          Logger.log(JSON.stringify(result.data, null, 2));
        }
      } else {
        Logger.log("No data returned");
      }
    } catch (error) {
      Logger.log("Error testing query: " + error);
    }
  }
  
  Logger.log("\n=== Test Complete ===");
}

/**
 * testSearchMeetings - פונקציית בדיקה לחיפוש פגישות
 */
function testSearchMeetings() {
  Logger.log("=== Starting Search Meetings Test ===");
  
  // בדיקה עם תאריכים של היום
  var today = new Date();
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  var params = {
    startDate: toIsoString(today),
    endDate: toIsoString(tomorrow)
  };
  
  Logger.log("Testing with params: " + JSON.stringify(params));
  
  try {
    var results = searchMeetings(params);
    Logger.log("Search completed successfully");
    Logger.log("Found " + results.length + " meetings");
    
    if (results.length > 0) {
      Logger.log("First meeting details:");
      Logger.log(JSON.stringify(results[0], null, 2));
    }
  } catch (error) {
    Logger.log("Error in search: " + error);
  }
  
  Logger.log("=== Test Complete ===");
}

/**
 * testChatbotWithLogging - פונקציית בדיקה מפורטת לצ'אטבוט
 */
function testChatbotWithLogging() {
  Logger.log("=== Starting Detailed Chatbot Test ===");
  
  var query = "כמה פגישות מתוכננות להיום?";
  Logger.log("\nTesting query: " + query);
  
  try {
    // 1. קריאה לפונקציית העיבוד
    Logger.log("\nStep 1: Calling processChatQuery");
    var result = processChatQuery(query);
    
    // 2. בדיקת הצלחה
    Logger.log("\nStep 2: Checking success");
    Logger.log("Success: " + result.success);
    
    // 3. בדיקת התשובה
    Logger.log("\nStep 3: Checking answer");
    Logger.log("Bot answer: " + result.answer);
    
    // 4. בדיקת הנתונים
    Logger.log("\nStep 4: Checking data");
    if (result.data) {
      if (Array.isArray(result.data)) {
        Logger.log("Found " + result.data.length + " meetings");
        
        if (result.data.length > 0) {
          Logger.log("\nFirst meeting details:");
          Logger.log(JSON.stringify(result.data[0], null, 2));
          
          // סיכום סטטוסים
          var statusCount = {};
          result.data.forEach(function(meeting) {
            var status = meeting.status || "לא מוגדר";
            statusCount[status] = (statusCount[status] || 0) + 1;
          });
          
          Logger.log("\nStatus breakdown:");
          Object.keys(statusCount).forEach(function(status) {
            Logger.log("- " + status + ": " + statusCount[status]);
          });
        }
      } else {
        Logger.log("Data is not an array:", result.data);
      }
    } else {
      Logger.log("No data returned");
    }
    
  } catch (error) {
    Logger.log("\nERROR in test:");
    Logger.log(error);
    Logger.log("Stack trace:", error.stack);
  }
  
  Logger.log("\n=== Test Complete ===");
}

/**
 * testChatbotJsonHandling - פונקציית בדיקה ספציפית לטיפול ב-JSON
 */
function testChatbotJsonHandling() {
  Logger.log("=== Starting JSON Handling Test ===");
  
  var testQueries = [
    "כמה פגישות יש היום?",
    "כמה מהפגישות של היום התקיימו?",
    "מי המדריכים שיש להם פגישות השבוע?"
  ];
  
  for (var i = 0; i < testQueries.length; i++) {
    var query = testQueries[i];
    Logger.log("\n=== Testing Query: " + query + " ===");
    
    try {
      // קריאה לפונקציית העיבוד
      var result = processChatQuery(query);
      
      // בדיקת הצלחה
      Logger.log("Success: " + result.success);
      
      if (result.success) {
        // בדיקת תקינות הנתונים
        if (result.data) {
          Logger.log("Data received successfully");
          Logger.log("Number of meetings: " + (Array.isArray(result.data) ? result.data.length : "N/A"));
          
          // הדפסת דוגמה לנתונים
          if (Array.isArray(result.data) && result.data.length > 0) {
            Logger.log("First meeting example:");
            Logger.log(JSON.stringify(result.data[0], null, 2));
          }
        } else {
          Logger.log("No data received");
        }
        
        // בדיקת התשובה
        Logger.log("\nBot response:");
        Logger.log(result.answer);
      } else {
        Logger.log("Error: " + result.answer);
      }
      
    } catch (error) {
      Logger.log("Test Error:");
      Logger.log(error);
      Logger.log("Stack trace:", error.stack);
    }
    
    Logger.log("\n---");
  }
  
  Logger.log("\n=== Test Complete ===");
}

/**
 * testFutureMeetingsSearch - פונקציית בדיקה מקיפה לחיפוש פגישות עתידיות
 * בודקת את כל השלבים בתהליך החיפוש ומדפיסה מידע מפורט
 */
function testFutureMeetingsSearch() {
  Logger.log("=== Starting Future Meetings Search Test ===");

  // 1. בדיקת חיפוש למחר
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  var tomorrowEnd = new Date(tomorrow);
  tomorrowEnd.setHours(23, 59, 59, 999);

  var startDateStr = toIsoString(tomorrow);
  var endDateStr = toIsoString(tomorrowEnd);

  Logger.log("\n1. Testing direct API query for tomorrow:");
  Logger.log("Start date: " + startDateStr);
  Logger.log("End date: " + endDateStr);

  // בדיקת קריאה ישירה ל-API
  var queryPayload = {
    objecttype: 6,
    page_size: 500,
    page_number: 1,
    fields: "activityid,scheduledstart,scheduledend,pcfsystemfield485,pcfsystemfield542,pcfsystemfield498,statuscode",
    query: "(scheduledstart >= " + startDateStr + ") AND " +
           "(scheduledstart < " + endDateStr + ") AND " +
           "(pcfsystemfield498 is-not-null)"
  };

  var directApiResult = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  Logger.log("Direct API Response:");
  Logger.log(JSON.stringify(directApiResult, null, 2));

  // 2. בדיקת החיפוש דרך searchMeetings
  Logger.log("\n2. Testing searchMeetings function:");
  var searchParams = {
    startDate: startDateStr,
    endDate: endDateStr
  };
  var searchResult = searchMeetings(searchParams);
  Logger.log("searchMeetings found " + (searchResult ? searchResult.length : 0) + " meetings");
  if (searchResult && searchResult.length > 0) {
    Logger.log("First meeting example:");
    Logger.log(JSON.stringify(searchResult[0], null, 2));
  }

  // 3. בדיקת החיפוש דרך הצ'אטבוט
  Logger.log("\n3. Testing chatbot query:");
  var chatQuery = "מה הפגישות שמתוכננות למחר?";
  var chatResult = processChatQuery(chatQuery);
  Logger.log("Chatbot analysis:");
  Logger.log(JSON.stringify(chatResult.analysis, null, 2));
  Logger.log("Meetings found: " + (chatResult.data ? chatResult.data.length : 0));

  // 4. בדיקת תנאי החיפוש בפונקציית getMeetingsForRange
  Logger.log("\n4. Testing getMeetingsForRange function:");
  var rangeResult = getMeetingsForRange(startDateStr, endDateStr);
  Logger.log("getMeetingsForRange found " + (rangeResult ? rangeResult.length : 0) + " meetings");
  
  // 5. בדיקת פגישות ללא סינון סטטוס
  Logger.log("\n5. Testing meetings without status filter:");
  var allMeetings = getMeetingsForRangeWithoutStatusFilter(startDateStr, endDateStr);
  Logger.log("Total meetings without status filter: " + (allMeetings ? allMeetings.length : 0));
  if (allMeetings && allMeetings.length > 0) {
    Logger.log("Status breakdown:");
    var statusCount = {};
    allMeetings.forEach(function(meeting) {
      var status = meeting.statuscode || "no_status";
      statusCount[status] = (statusCount[status] || 0) + 1;
    });
    Logger.log(JSON.stringify(statusCount, null, 2));
  }

  // 6. בדיקת תקינות התאריכים
  Logger.log("\n6. Date format validation:");
  Logger.log("Tomorrow date object: " + tomorrow);
  Logger.log("Tomorrow ISO string: " + startDateStr);
  Logger.log("End date object: " + tomorrowEnd);
  Logger.log("End ISO string: " + endDateStr);

  // 7. סיכום ממצאים
  Logger.log("\n=== Summary of Findings ===");
  Logger.log("Direct API query found: " + (directApiResult && directApiResult.data ? directApiResult.data.Data.length : 0) + " meetings");
  Logger.log("searchMeetings found: " + (searchResult ? searchResult.length : 0) + " meetings");
  Logger.log("Chatbot query found: " + (chatResult.data ? chatResult.data.length : 0) + " meetings");
  Logger.log("getMeetingsForRange found: " + (rangeResult ? rangeResult.length : 0) + " meetings");
  Logger.log("Without status filter found: " + (allMeetings ? allMeetings.length : 0) + " meetings");

  Logger.log("\n=== Test Complete ===");
  return {
    directApiResults: directApiResult,
    searchResults: searchResult,
    chatbotResults: chatResult,
    rangeResults: rangeResult,
    allMeetingsResults: allMeetings
  };
}

/**
 * testSpecificDateRange - בודק טווח תאריכים ספציפי
 * @param {string} startDate - תאריך התחלה בפורמט YYYY-MM-DD
 * @param {string} endDate - תאריך סיום בפורמט YYYY-MM-DD
 */
function testSpecificDateRange(startDate, endDate) {
  Logger.log("=== Testing Specific Date Range ===");
  Logger.log("Start Date: " + startDate);
  Logger.log("End Date: " + endDate);

  var startDateObj = new Date(startDate);
  var endDateObj = new Date(endDate);
  startDateObj.setHours(0, 0, 0, 0);
  endDateObj.setHours(23, 59, 59, 999);

  var startDateStr = toIsoString(startDateObj);
  var endDateStr = toIsoString(endDateObj);

  var result = testFutureMeetingsSearch();
  
  // הדפסת תוצאות מפורטות
  Logger.log("\n=== Detailed Results for " + startDate + " to " + endDate + " ===");
  if (result.allMeetingsResults && result.allMeetingsResults.length > 0) {
    Logger.log("\nMeetings found:");
    result.allMeetingsResults.forEach(function(meeting, index) {
      Logger.log("\nMeeting " + (index + 1) + ":");
      Logger.log("Start time: " + meeting.scheduledstart);
      Logger.log("Status: " + meeting.statuscode);
      Logger.log("Cycle ID: " + meeting.pcfsystemfield498);
    });
  } else {
    Logger.log("No meetings found in this date range");
  }

  return result;
}

/**
 * ChatMemory - מערכת זיכרון לצ'אטבוט
 * מאפשרת שמירה ושליפה של מידע היסטורי על שיחות ופגישות
 */

const CHAT_MEMORY_KEY = "CHAT_MEMORY";
const MAX_MEMORY_ITEMS = 100;  // מספר מקסימלי של פריטים בזיכרון

/**
 * saveToMemory - שומר מידע בזיכרון הצ'אטבוט
 * @param {object} data - המידע לשמירה
 * @param {string} type - סוג המידע (meeting/conversation/status)
 * @param {string} key - מזהה ייחודי למידע
 */
function saveToMemory(data, type, key) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let memory = JSON.parse(scriptProperties.getProperty(CHAT_MEMORY_KEY) || "{}");
    
    if (!memory[type]) {
      memory[type] = {};
    }
    
    // הוספת חותמת זמן לנתונים
    const memoryItem = {
      data: data,
      timestamp: new Date().toISOString(),
      key: key
    };
    
    memory[type][key] = memoryItem;
    
    // מחיקת פריטים ישנים אם עברנו את המקסימום
    const typeItems = Object.values(memory[type]);
    if (typeItems.length > MAX_MEMORY_ITEMS) {
      // מיון לפי זמן ומחיקת הפריטים הישנים ביותר
      typeItems.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      memory[type] = {};
      typeItems.slice(0, MAX_MEMORY_ITEMS).forEach(item => {
        memory[type][item.key] = item;
      });
    }
    
    scriptProperties.setProperty(CHAT_MEMORY_KEY, JSON.stringify(memory));
    Logger.log(`Saved to memory - Type: ${type}, Key: ${key}`);
    return true;
  } catch (error) {
    Logger.log(`Error saving to memory: ${error}`);
    return false;
  }
}

/**
 * getFromMemory - מחזיר מידע מהזיכרון
 * @param {string} type - סוג המידע
 * @param {string} key - מזהה המידע
 * @return {object} המידע השמור או null אם לא נמצא
 */
function getFromMemory(type, key) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const memory = JSON.parse(scriptProperties.getProperty(CHAT_MEMORY_KEY) || "{}");
    
    if (memory[type] && memory[type][key]) {
      Logger.log(`Retrieved from memory - Type: ${type}, Key: ${key}`);
      return memory[type][key].data;
    }
    
    return null;
  } catch (error) {
    Logger.log(`Error retrieving from memory: ${error}`);
    return null;
  }
}

/**
 * searchMemory - חיפוש בזיכרון לפי סוג וקריטריונים
 * @param {string} type - סוג המידע לחיפוש
 * @param {function} filterFn - פונקציית סינון
 * @return {array} מערך של תוצאות מתאימות
 */
function searchMemory(type, filterFn) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const memory = JSON.parse(scriptProperties.getProperty(CHAT_MEMORY_KEY) || "{}");
    
    if (!memory[type]) {
      return [];
    }
    
    const results = Object.values(memory[type])
      .filter(item => filterFn(item.data))
      .map(item => item.data);
    
    Logger.log(`Search memory - Type: ${type}, Results: ${results.length}`);
    return results;
  } catch (error) {
    Logger.log(`Error searching memory: ${error}`);
    return [];
  }
}

/**
 * clearMemoryType - מוחק את כל המידע מסוג מסוים
 * @param {string} type - סוג המידע למחיקה
 */
function clearMemoryType(type) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let memory = JSON.parse(scriptProperties.getProperty(CHAT_MEMORY_KEY) || "{}");
    
    if (memory[type]) {
      delete memory[type];
      scriptProperties.setProperty(CHAT_MEMORY_KEY, JSON.stringify(memory));
      Logger.log(`Cleared memory type: ${type}`);
    }
  } catch (error) {
    Logger.log(`Error clearing memory type: ${error}`);
  }
}

/**
 * testChatMemory - פונקציית בדיקה למערכת הזיכרון
 */
function testChatMemory() {
  Logger.log("=== Testing Chat Memory System ===");
  
  // 1. בדיקת שמירה וקריאה בסיסית
  const testData = {
    id: "test123",
    startDate: "2024-03-19T10:00:00",
    status: "התקיימה",
    description: "פגישת בדיקה"
  };
  
  Logger.log("\n1. Testing basic save and retrieve:");
  const saved = saveToMemory(testData, 'meetings', 'test123');
  Logger.log("Save result: " + saved);
  
  const retrieved = getFromMemory('meetings', 'test123');
  Logger.log("Retrieved data: " + JSON.stringify(retrieved));
  
  // 2. בדיקת חיפוש
  Logger.log("\n2. Testing search functionality:");
  const searchResults = searchMemory('meetings', 
    meeting => meeting.status === "התקיימה"
  );
  Logger.log("Search results: " + JSON.stringify(searchResults));
  
  // 3. בדיקת ניקוי
  Logger.log("\n3. Testing memory cleanup:");
  clearMemoryType('meetings');
  const afterClear = getFromMemory('meetings', 'test123');
  Logger.log("After clear: " + JSON.stringify(afterClear));
  
  Logger.log("\n=== Memory Test Complete ===");
}

/**
 * getInstitutionalOrders - שליפת הזמנות מוסדיות בסטטוס מעקב וגבייה והסתיימה
 * @param {string} startDate - תאריך התחלה בפורמט YYYY-MM-DD (אופציונלי)
 * @param {string} endDate - תאריך סיום בפורמט YYYY-MM-DD (אופציונלי)
 * @return {Array} מערך של הזמנות מוסדיות עם פרטיהן
 */
function getInstitutionalOrders(startDate, endDate) {
  Logger.log("getInstitutionalOrders called with date range:", startDate, "to", endDate);
  
  var query = [
    "(pcfsystemfield182 = 10) OR (pcfsystemfield182 = 13)"  // סינון לפי סטטוס מעקב וגבייה (10) או הסתיים (13)
  ];

  // הוספת סינון לפי טווח תאריכים אם סופקו
  if (startDate) {
    query.push("(pcfsystemfield197 >= " + startDate + "T00:00:00)");
  }
  if (endDate) {
    query.push("(pcfsystemfield197 <= " + endDate + "T23:59:59)");
  }

  var queryPayload = {
    objecttype: 1005,  // ObjectType עבור הזמנות מוסדיות
    page_size: 200,
    page_number: 1,
    fields: "customobject1005id,name,pcfsystemfield192,pcfsystemfield182,pcfsystemfield181,pcfsystemfield197",  // הוספת שדה תאריך תחילת פעילות
    query: query.join(" AND ")
  };

  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("Error fetching institutional orders");
    return null;
  }

  return data.data.Data;
}

// ... existing code ...

/**
 * printInstitutionalOrders - מחזירה את ההזמנות המוסדיות בפורמט מתאים לתצוגה
 * @param {string} startDate - תאריך התחלה בפורמט YYYY-MM-DD (אופציונלי)
 * @param {string} endDate - תאריך סיום בפורמט YYYY-MM-DD (אופציונלי)
 * @return {Array} מערך של הזמנות מוסדיות מעובדות לתצוגה
 */
function printInstitutionalOrders(startDate, endDate) {
  const orders = getInstitutionalOrders(startDate, endDate);
  if (!orders || orders.length === 0) {
    return [];
  }
  
  return orders.map((order) => {
    const payment = parseFloat(order.pcfsystemfield192 || 0);
    const statusNum = order.pcfsystemfield182;
    let statusText = "";
    
    // המרה ממספר לטקסט
    if (statusNum === "10") {
      statusText = "מעקב וגבייה";
    } else if (statusNum === "13") {
      statusText = "הסתיים";
    } else {
      statusText = statusNum;
    }

    // המרת תאריך הפעילות לפורמט מקומי
    const activityDate = order.pcfsystemfield197 ? 
      new Date(order.pcfsystemfield197).toLocaleDateString('he-IL') : 'לא צוין';

    return {
      Name: order.name || '',
      OrderId: order.customobject1005id,
      TotalPayment: payment.toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      BranchName: getBranchName(order.pcfsystemfield181),
      Status: statusText,
      ActivityStartDate: activityDate,
      RawStatus: statusNum
    };
  });
}

/**
 * testInstitutionalOrdersWithDateRange - פונקציית בדיקה להזמנות מוסדיות עם טווח תאריכים
 * @param {string} startDate - תאריך התחלה בפורמט YYYY-MM-DD
 * @param {string} endDate - תאריך סיום בפורמט YYYY-MM-DD
 */
function testInstitutionalOrdersWithDateRange(startDate, endDate) {
  Logger.log("=== Testing Institutional Orders with Date Range ===");
  Logger.log("Date Range:", startDate, "to", endDate);
  
  const orders = getInstitutionalOrders(startDate, endDate);
  Logger.log("Raw orders from API:", orders);

  const processedOrders = printInstitutionalOrders(startDate, endDate);
  Logger.log("Processed orders:", processedOrders);

  // חישוב סכומים לפי סטטוס
  let totalTracking = 0;
  let totalCompleted = 0;
  
  processedOrders.forEach(order => {
    const amount = parseFloat(order.TotalPayment.replace(/[₪,\s]/g, ''));
    Logger.log("Processing order:", {
      name: order.Name,
      status: order.Status,
      activityStartDate: order.ActivityStartDate,
      amount: amount
    });

    if (order.RawStatus === "10") {
      totalTracking += amount;
    } else if (order.RawStatus === "13") {
      totalCompleted += amount;
    }
  });

  Logger.log("Summary:", {
    totalOrders: processedOrders.length,
    tracking: totalTracking.toLocaleString('he-IL'),
    completed: totalCompleted.toLocaleString('he-IL'),
    total: (totalTracking + totalCompleted).toLocaleString('he-IL')
  });
}

/**
 * getInstitutionalOrdersByAcademicYear - שליפת הזמנות מוסדיות לפי שנת פעילות
 * @param {number} year - שנת הפעילות (לדוגמה: 2024 עבור שנת 2024-2025)
 * @return {Array} מערך של הזמנות מוסדיות עם פרטיהן
 */
function getInstitutionalOrdersByAcademicYear(year) {
  Logger.log("getInstitutionalOrdersByAcademicYear called for year:", year);
  
  // חישוב טווח התאריכים לשנת הפעילות
  var startDate = year + "-09-01";          // 1 בספטמבר של השנה הנוכחית
  var endDate = (year + 1) + "-06-30";      // 30 ביוני של השנה העוקבת

  var query = [
    "(pcfsystemfield182 = 10) OR (pcfsystemfield182 = 13)",  // סינון לפי סטטוס מעקב וגבייה (10) או הסתיים (13)
    "(pcfsystemfield197 >= " + startDate + "T00:00:00)",     // תאריך תחילת פעילות מ-1 בספטמבר
    "(pcfsystemfield197 <= " + endDate + "T23:59:59)"        // תאריך תחילת פעילות עד 30 ביוני
  ];

  var queryPayload = {
    objecttype: 1005,
    page_size: 200,
    page_number: 1,
    fields: "customobject1005id,name,pcfsystemfield192,pcfsystemfield182,pcfsystemfield181,pcfsystemfield197",
    query: query.join(" AND ")
  };

  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("Error fetching institutional orders");
    return null;
  }

  return data.data.Data;
}

/**
 * printInstitutionalOrdersByAcademicYear - מחזירה את ההזמנות המוסדיות בפורמט מתאים לתצוגה לפי שנת פעילות
 * @param {number} year - שנת הפעילות (לדוגמה: 2024 עבור שנת 2024-2025)
 * @return {Array} מערך של הזמנות מוסדיות מעובדות לתצוגה
 */
function printInstitutionalOrdersByAcademicYear(year) {
  const orders = getInstitutionalOrdersByAcademicYear(year);
  if (!orders || orders.length === 0) {
    return [];
  }
  
  return orders.map((order) => {
    const payment = parseFloat(order.pcfsystemfield192 || 0);
    const statusNum = order.pcfsystemfield182;
    let statusText = "";
    
    // המרה ממספר לטקסט
    if (statusNum === "10") {
      statusText = "מעקב וגבייה";
    } else if (statusNum === "13") {
      statusText = "הסתיים";
    } else {
      statusText = statusNum;
    }

    // המרת תאריך הפעילות לפורמט מקומי
    const activityDate = order.pcfsystemfield197 ? 
      new Date(order.pcfsystemfield197).toLocaleDateString('he-IL') : 'לא צוין';

    return {
      Name: order.name || '',
      OrderId: order.customobject1005id,
      TotalPayment: payment.toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      BranchName: getBranchName(order.pcfsystemfield181),
      Status: statusText,
      ActivityStartDate: activityDate,
      RawStatus: statusNum,
      AcademicYear: `${year}-${year + 1}`
    };
  });
}

/**
 * getCurrentAcademicYear - מחזיר את שנת הפעילות הנוכחית
 * @return {number} שנת הפעילות הנוכחית
 */
function getCurrentAcademicYear() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // getMonth() מחזיר 0-11
  
  // אם אנחנו בחודשים יולי או אוגוסט, נחשיב את זה כשנת הפעילות הבאה
  if (month >= 7) {
    return year;
  }
  return year - 1;
}

/**
 * testInstitutionalOrdersForCurrentYear - פונקציית בדיקה להזמנות מוסדיות בשנת הפעילות הנוכחית
 */
function testInstitutionalOrdersForCurrentYear() {
  const currentYear = getCurrentAcademicYear();
  Logger.log(`=== Testing Institutional Orders for Academic Year ${currentYear}-${currentYear + 1} ===`);
  
  const orders = getInstitutionalOrdersByAcademicYear(currentYear);
  Logger.log("Raw orders from API:", orders);

  const processedOrders = printInstitutionalOrdersByAcademicYear(currentYear);
  Logger.log("Processed orders:", processedOrders);

  // חישוב סכומים לפי סטטוס
  let totalTracking = 0;
  let totalCompleted = 0;
  
  processedOrders.forEach(order => {
    const amount = parseFloat(order.TotalPayment.replace(/[₪,\s]/g, ''));
    Logger.log("Processing order:", {
      name: order.Name,
      status: order.Status,
      activityStartDate: order.ActivityStartDate,
      amount: amount
    });

    if (order.RawStatus === "10") {
      totalTracking += amount;
    } else if (order.RawStatus === "13") {
      totalCompleted += amount;
    }
  });

  Logger.log("Summary:", {
    academicYear: `${currentYear}-${currentYear + 1}`,
    totalOrders: processedOrders.length,
    tracking: totalTracking.toLocaleString('he-IL'),
    completed: totalCompleted.toLocaleString('he-IL'),
    total: (totalTracking + totalCompleted).toLocaleString('he-IL')
  });
}

/**
 * getDigitalCourseRegistrations - מחזיר את כל ההרשמות לקורסים דיגיטליים
 * @return {Array} מערך של הרשמות לקורסים דיגיטליים
 */
function getDigitalCourseRegistrations() {
  Logger.log("getDigitalCourseRegistrations called");
  
  var queryPayload = {
    objecttype: 33,  // אובייקט הרשמות
    page_size: 200,
    page_number: 1,
    fields: "accountproductid,accountid,pcfsystemfield129,pcfsystemfield53,statuscode",  // מזהה, שם הרוכש, תאריך הרשמה, קורס, סטטוס
    query: "statuscode = 8"  // סטטוס נרשם = 8
  };

  var registrations = [];
  var pageNumber = 1;
  var maxPages = 10;

  while (true) {
    Logger.log("מושך עמוד " + pageNumber + " של הרשמות");
    queryPayload.page_number = pageNumber;
    var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
    
    if (!data || !data.data || !data.data.Data) {
      Logger.log("שגיאה בקבלת נתוני הרשמות בעמוד " + pageNumber);
      break;
    }

    var currentRegistrations = data.data.Data;
    Logger.log("התקבלו " + currentRegistrations.length + " הרשמות בעמוד " + pageNumber);
    
    // סינון רק הרשמות לקורסים דיגיטליים
    for (var i = 0; i < currentRegistrations.length; i++) {
      var reg = currentRegistrations[i];
      var courseId = reg.pcfsystemfield53;
      if (!courseId) {
        Logger.log("נמצאה הרשמה ללא מזהה קורס:", reg.accountproductid);
        continue;
      }

      // בדיקה אם הקורס הוא דיגיטלי
      var course = getCourseById(courseId);
      if (!course) {
        Logger.log("לא נמצא קורס עם מזהה:", courseId);
        continue;
      }

      Logger.log("בודק קורס:", {
        id: courseId,
        name: course.name,
        type: course.pcfsystemfield44
      });

      if (course.pcfsystemfield44 === "2") { // קורס דיגיטלי
        Logger.log("נמצא קורס דיגיטלי:", course.name);
        registrations.push(reg);
      }
    }

    if (currentRegistrations.length < queryPayload.page_size || pageNumber >= maxPages) {
      break;
    }
    
    pageNumber++;
    Utilities.sleep(300);
  }

  Logger.log("סך הכל נמצאו " + registrations.length + " הרשמות לקורסים דיגיטליים");
  return registrations;
}

/**
 * getCourseById - מחזיר פרטי קורס לפי מזהה
 * @param {string} courseId - מזהה הקורס
 * @return {object} פרטי הקורס
 */
function getCourseById(courseId) {
  if (!courseId) return null;
  
  var url = API_URL + "/14/" + courseId;  // אובייקט קורסים = 14
  var data = sendGetWithRetry(url, MAX_RETRIES, RETRY_DELAY_MS);
  
  if (!data || !data.data || !data.data.Record) {
    Logger.log("Failed to fetch course: " + courseId);
    return null;
  }
  
  return data.data.Record;
}

/**
 * printDigitalCourseRegistrations - מחזיר את ההרשמות בפורמט מתאים לתצוגה
 * @return {Array} מערך של הרשמות מעובדות לתצוגה
 */
function printDigitalCourseRegistrations() {
  const registrations = getDigitalCourseRegistrations();
  if (!registrations || registrations.length === 0) {
    return [];
  }
  
  return registrations.map((reg) => {
    const course = getCourseById(reg.pcfsystemfield53);
    const registrationDate = reg.pcfsystemfield129 ? 
      new Date(reg.pcfsystemfield129).toLocaleDateString('he-IL') : 'לא צוין';

    return {
      RegistrationId: reg.accountproductid,
      BuyerName: reg.accountid || 'לא ידוע',
      CourseName: course ? course.name : 'לא ידוע',
      RegistrationDate: registrationDate,
      CourseId: reg.pcfsystemfield53
    };
  });
}

/**
 * testDigitalCourseRegistrations - פונקציית בדיקה להרשמות לקורסים דיגיטליים
 * בודקת את הפרמטרים של הקריאה ומדפיסה את התוצאות לבדיקה
 */
function testDigitalCourseRegistrations() {
  Logger.log("=== בדיקת הרשמות לקורסים דיגיטליים ===");
  
  // שליפת כל ההרשמות
  var queryPayload = {
    objecttype: 33,  // אובייקט הרשמות
    page_size: 200,
    page_number: 1,
    fields: "accountproductid,accountid,pcfsystemfield129,pcfsystemfield53,statuscode",
    query: "statuscode = 8"  // סטטוס נרשם = 8
  };
  
  Logger.log("שולח בקשה עם הפרמטרים:", queryPayload);
  
  var data = sendRequestWithRetry(QUERY_API_URL, queryPayload, MAX_RETRIES, RETRY_DELAY_MS);
  if (!data || !data.data || !data.data.Data) {
    Logger.log("שגיאה בקבלת נתוני הרשמות:", data);
    return;
  }
  
  Logger.log("מספר הרשמות שהתקבלו:", data.data.Data.length);
  
  // בדיקת כל הרשמה
  data.data.Data.forEach(function(registration) {
    Logger.log("\nבדיקת הרשמה:", registration.accountproductid);
    Logger.log("שם הרוכש:", registration.accountid);
    
    // שליפת פרטי הקורס
    var courseId = registration.pcfsystemfield53;
    Logger.log("מזהה קורס:", courseId);
    
    var course = getCourseById(courseId);
    if (course) {
      Logger.log("פרטי הקורס:", {
        id: course.customobject14id,
        name: course.name,
        type: course.pcfsystemfield44  // סוג הקורס
      });
    } else {
      Logger.log("לא נמצא קורס עם מזהה:", courseId);
    }
    
    Logger.log("תאריך הרשמה:", registration.pcfsystemfield129);
    Logger.log("סטטוס:", registration.statuscode);
  });
}
