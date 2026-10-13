const SHEET_ID = "1X-vkKWlFoFcYS9Vr6ofafE8O3TqHhl9OEjbmnc1U_PI";
const DRIVE_FOLDER_ID = "PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE";
const SHEET_NAME = "Poster Submissions";
const MAX_FILE_BYTES = 8 * 1024 * 1024;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const required = [
      "presenterName", "email", "mobile", "institution",
      "department", "category", "posterTitle", "track", "abstract"
    ];

    required.forEach(key => {
      if (!String(data[key] || "").trim()) {
        throw new Error("Missing field: " + key);
      }
    });

    if (data.submissionType !== "poster") {
      throw new Error("Invalid submission type.");
    }

    if (!data.originalityConfirmed || !data.policyAccepted) {
      throw new Error("Please accept the required declarations.");
    }

    const file = data.posterFile;

    if (!file || !file.data || !/\.pdf$/i.test(file.name)) {
      throw new Error("Please upload a PDF poster.");
    }

    const bytes = Utilities.base64Decode(file.data);

    if (bytes.length > MAX_FILE_BYTES) {
      throw new Error("PDF must be 8 MB or smaller.");
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Timestamp", "Poster ID", "Presenter Name", "Email",
        "Mobile", "Institution", "Department", "Category",
        "Poster Title", "Track", "Co-authors", "Abstract",
        "Poster PDF URL", "Status"
      ]);
      sheet.setFrozenRows(1);
    }

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const posterId =
      "RATE2026-POST-" +
      Utilities.getUuid().slice(0, 8).toUpperCase();

    const safeName = file.name.replace(/[^\w.\- ]/g, "_");
    const blob = Utilities.newBlob(
      bytes,
      "application/pdf",
      posterId + "_" + safeName
    );

    const driveFile = folder.createFile(blob);

    sheet.appendRow([
      new Date(),
      posterId,
      data.presenterName,
      data.email,
      data.mobile,
      data.institution,
      data.department,
      data.category,
      data.posterTitle,
      data.track,
      data.coAuthors || "",
      data.abstract,
      driveFile.getUrl(),
      "Submitted"
    ]);

    return json({
      success: true,
      posterId: posterId
    });

  } catch (error) {
    return json({
      success: false,
      error: error.message || "Submission failed."
    });
  }
}

function doGet() {
  return json({
    service: "IC-RATE 2026 Poster Submission",
    status: "ready"
  });
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
