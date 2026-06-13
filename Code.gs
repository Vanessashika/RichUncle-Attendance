/**
 * RICHUNCLE ATTENDANCE — Apps Script Backend
 *
 * SETUP:
 * 1. Go to sheets.google.com -> create a new blank spreadsheet, name it "Richuncle Attendance".
 * 2. Rename "Sheet1" tab to "Attendance".
 * 3. In row 1, add these headers (columns A-H):
 *    Timestamp | Name | Latitude | Longitude | Distance (m) | Status | Location | IP/Device
 * 4. Extensions -> Apps Script. Delete any starter code and paste this whole file in.
 * 5. Update the LOCATIONS array and ALLOWED_RADIUS_M below if needed (already set for Richuncle's two shops).
 * 6. Click Deploy -> New deployment -> Type: "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Click Deploy. Copy the Web App URL — this goes into the attendance page's SCRIPT_URL.
 * 8. Re-deploy (Deploy -> Manage deployments -> edit -> new version) any time you change this script.
 */

const LOCATIONS = [
  { name: "Shop A", lat: 5.548706, lng: -0.207453 },
  { name: "Shop B", lat: 5.550745, lng: -0.205991 }
];
const ALLOWED_RADIUS_M = 100;
const SHEET_NAME = "Attendance";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const name = (data.name || "").toString().trim();
    const lat = parseFloat(data.lat);
    const lng = parseFloat(data.lng);

    if (!name) {
      return jsonResponse({ success: false, error: "Name is required." });
    }
    if (isNaN(lat) || isNaN(lng)) {
      return jsonResponse({ success: false, error: "Location data missing or invalid." });
    }

    let best = null;
    for (const loc of LOCATIONS) {
      const d = haversineDistance(loc.lat, loc.lng, lat, lng);
      if (best === null || d < best.distance) {
        best = { name: loc.name, distance: d };
      }
    }

    const withinRange = best.distance <= ALLOWED_RADIUS_M;
    const status = withinRange ? "Within Range" : "Out of Range";
    const matchedLocation = withinRange ? best.name : "—";

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const now = new Date();
    sheet.appendRow([
      now,
      name,
      lat,
      lng,
      Math.round(best.distance),
      status,
      matchedLocation,
      data.device || ""
    ]);

    if (!withinRange) {
      return jsonResponse({
        success: false,
        error: "You are outside all allowed locations (closest: " + best.name + ", " + Math.round(best.distance) + "m away, limit is " + ALLOWED_RADIUS_M + "m). Attendance not counted.",
        distance: Math.round(best.distance)
      });
    }

    return jsonResponse({
      success: true,
      message: "Attendance marked for " + name + " at " + best.name + ".",
      location: best.name,
      distance: Math.round(best.distance),
      timestamp: now.toISOString()
    });

  } catch (err) {
    return jsonResponse({ success: false, error: "Server error: " + err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "Richuncle Attendance API is running." });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Haversine formula — distance between two lat/lng points in meters.
 */
function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Earth radius in meters
  const toRad = (deg) => (deg * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
