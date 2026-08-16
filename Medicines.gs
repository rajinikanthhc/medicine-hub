/* =========================================
   MEDICINES
========================================= */

const MEDICINE_SPREADSHEET_ID =
  "1Lpm8iW4RNnkvw4MiqpzfGi-H5Uh9dTIRztARmE4ZpJg";

const MEDICINE_SHEET_NAME = "Medicines";


/* =========================================
   GET SHEET
========================================= */

function getMedicineSheet() {

  const ss =
    SpreadsheetApp.openById(
      MEDICINE_SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      MEDICINE_SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      'Sheet "Medicines" not found.'
    );
  }

  return sheet;
}


/* =========================================
   GET ALL MEDICINES
========================================= */

function getMedicines() {

  const sheet =
    getMedicineSheet();

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return [];
  }


  /*
   * Read 9 columns
   *
   * A = ID
   * B = Name
   * C = Generic Name
   * D = Used For
   * E = Use
   * F = Description
   * G = Photo
   * H = Favorite
   * I = Category
   */

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        9
      )
      .getValues();


  return values

    .filter(function (row) {

      return row[0] !== "" &&
             row[1] !== "";

    })

    .map(function (row) {

      return {

        id:
          row[0],

        name:
          row[1],

        genericName:
          row[2],

        usedFor:
          row[3],

        use:
          row[4],

        description:
          row[5],

        photo:
          row[6],

        favorite:
          row[7] === true ||
          String(row[7])
            .toLowerCase() === "true" ||
          String(row[7]) === "1" ||
          String(row[7])
            .toLowerCase() === "yes",

        category:
          row[8] || ""

      };

    });

}


/* =========================================
   ADD MEDICINE
========================================= */

function addMedicine(medicine) {

  if (!medicine) {

    throw new Error(
      "Medicine data is missing."
    );

  }


  const sheet =
    getMedicineSheet();


  /*
   * DUPLICATE CHECK
   */

  const lastRow =
    sheet.getLastRow();


  if (lastRow >= 2) {

    const existingNames =
      sheet
        .getRange(
          2,
          2,
          lastRow - 1,
          1
        )
        .getValues();


    const newName =
      String(
        medicine.name || ""
      )
        .trim()
        .toLowerCase();


    for (
      let i = 0;
      i < existingNames.length;
      i++
    ) {

      const existingName =
        String(
          existingNames[i][0] || ""
        )
          .trim()
          .toLowerCase();


      if (
        existingName &&
        existingName === newName
      ) {

        throw new Error(
          "Duplicate medicine: " +
          existingNames[i][0]
        );

      }

    }

  }


  const newId =
    getLowestAvailableMedicineId(
      sheet
    );


  sheet.appendRow([

    newId,

    medicine.name || "",

    medicine.genericName || "",

    medicine.usedFor || "",

    medicine.use || "",

    medicine.description || "",

    medicine.photo || "",

    medicine.favorite === true
      ? true
      : false,

    medicine.category || ""

  ]);


  return {

    success: true,

    id: newId

  };

}


/* =========================================
   FIND LOWEST AVAILABLE ID
========================================= */

function getLowestAvailableMedicineId(
  sheet
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return 1;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  const usedIds = {};


  values.forEach(function (row) {

    const id =
      Number(row[0]);


    if (
      Number.isInteger(id) &&
      id > 0
    ) {

      usedIds[id] = true;

    }

  });


  let id = 1;


  while (usedIds[id]) {
    id++;
  }


  return id;

}


/* =========================================
   UPDATE MEDICINE
========================================= */

function updateMedicine(medicine) {

  if (!medicine) {

    throw new Error(
      "Medicine data is missing."
    );

  }


  if (!medicine.id) {

    throw new Error(
      "Medicine ID is missing."
    );

  }


  const sheet =
    getMedicineSheet();


  const rowNumber =
    findMedicineRow(
      sheet,
      medicine.id
    );


  if (rowNumber === -1) {

    throw new Error(
      "Medicine not found."
    );

  }


  /*
   * DUPLICATE CHECK
   */

  const lastRow =
    sheet.getLastRow();


  const newName =
    String(
      medicine.name || ""
    )
      .trim()
      .toLowerCase();


  if (lastRow >= 2) {

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          2
        )
        .getValues();


    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      const existingId =
        Number(values[i][0]);


      const existingName =
        String(
          values[i][1] || ""
        )
          .trim()
          .toLowerCase();


      if (
        existingId !== Number(medicine.id) &&
        existingName === newName
      ) {

        throw new Error(
          "Duplicate medicine: " +
          values[i][1]
        );

      }

    }

  }


  /*
   * Keep existing photo if
   * no new photo was supplied.
   */

  const oldPhoto =
    sheet
      .getRange(
        rowNumber,
        7
      )
      .getValue();


  const photo =
    medicine.photo
      ? medicine.photo
      : oldPhoto;


  /*
   * Keep existing favorite if
   * not supplied.
   */

  const oldFavorite =
    sheet
      .getRange(
        rowNumber,
        8
      )
      .getValue();


  const favorite =
    typeof medicine.favorite ===
    "boolean"

      ? medicine.favorite

      : (
          oldFavorite === true ||
          String(oldFavorite)
            .toLowerCase() === "true" ||
          String(oldFavorite) === "1"
        );


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      9
    )
    .setValues([

      [

        Number(medicine.id),

        medicine.name || "",

        medicine.genericName || "",

        medicine.usedFor || "",

        medicine.use || "",

        medicine.description || "",

        photo || "",

        favorite,

        medicine.category || ""

      ]

    ]);


  return {

    success: true

  };

}


/* =========================================
   TOGGLE FAVORITE
========================================= */

function toggleFavorite(id) {

  const sheet =
    getMedicineSheet();


  const rowNumber =
    findMedicineRow(
      sheet,
      id
    );


  if (rowNumber === -1) {

    throw new Error(
      "Medicine not found."
    );

  }


  const currentValue =
    sheet
      .getRange(
        rowNumber,
        8
      )
      .getValue();


  const currentFavorite =
    currentValue === true ||
    String(currentValue)
      .toLowerCase() === "true" ||
    String(currentValue) === "1" ||
    String(currentValue)
      .toLowerCase() === "yes";


  const newFavorite =
    !currentFavorite;


  sheet
    .getRange(
      rowNumber,
      8
    )
    .setValue(
      newFavorite
    );


  return {

    success: true,

    favorite:
      newFavorite

  };

}


/* =========================================
   FIND MEDICINE ROW
========================================= */

function findMedicineRow(
  sheet,
  medicineId
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return -1;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  const id =
    Number(medicineId);


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      Number(values[i][0]) === id
    ) {

      return i + 2;

    }

  }


  return -1;

}


/* =========================================
   DELETE MEDICINE
========================================= */

function deleteMedicine(
  id,
  passcode
) {

  if (
    String(passcode) !== "12345"
  ) {

    throw new Error(
      "Incorrect passcode."
    );

  }


  const sheet =
    getMedicineSheet();


  const rowNumber =
    findMedicineRow(
      sheet,
      id
    );


  if (rowNumber === -1) {

    throw new Error(
      "Medicine not found."
    );

  }


  sheet.deleteRow(
    rowNumber
  );


  return {

    success: true

  };

}


/* =========================================
   UPLOAD MEDICINE PHOTO TO GITHUB
========================================= */

function uploadMedicinePhoto(
  fileData,
  fileName,
  medicineName
) {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "GITHUB_TOKEN"
      );


  if (!token) {

    throw new Error(
      "GITHUB_TOKEN not found in Script Properties."
    );

  }


  const owner =
    "rajinikanthhc";

  const repo =
    "images";

  const branch =
    "main";

  const folder =
    "medicines";


  if (
    !fileData ||
    !fileName
  ) {

    throw new Error(
      "Photo file is missing."
    );

  }


  if (!medicineName) {

    throw new Error(
      "Medicine name is missing."
    );

  }


  const base64Data =
    fileData.indexOf(",") >= 0

      ? fileData.split(",")[1]

      : fileData;


  /* =========================================
     CREATE CLEAN FILE NAME
  ========================================= */

  const cleanName =
    medicineName
      .trim()
      .replace(
        /[^a-zA-Z0-9]+/g,
        "_"
      )
      .replace(
        /^_+|_+$/g,
        "");


  const originalExtension =
    fileName.includes(".")

      ? fileName.substring(
          fileName.lastIndexOf(".")
        )

      : ".jpg";


  const finalFileName =
    cleanName +
    originalExtension.toLowerCase();


  const path =
    folder +
    "/" +
    finalFileName;


  const apiUrl =
    "https://api.github.com/repos/" +
    owner +
    "/" +
    repo +
    "/contents/" +
    path;


  /* =========================================
     CHECK EXISTING FILE
  ========================================= */

  let existingSha = null;


  const checkResponse =
    UrlFetchApp.fetch(
      apiUrl,
      {

        method: "get",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        muteHttpExceptions:
          true

      }
    );


  if (
    checkResponse
      .getResponseCode() === 200
  ) {

    const existing =
      JSON.parse(
        checkResponse
          .getContentText()
      );


    existingSha =
      existing.sha;

  }


  /* =========================================
     UPLOAD / REPLACE
  ========================================= */

  const payload = {

    message:
      existingSha

        ? "Update medicine image: " +
          finalFileName

        : "Add medicine image: " +
          finalFileName,

    content:
      base64Data,

    branch:
      branch

  };


  if (existingSha) {

    payload.sha =
      existingSha;

  }


  const response =
    UrlFetchApp.fetch(
      apiUrl,
      {

        method: "put",

        contentType:
          "application/json",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true

      }
    );


  const status =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  if (
    status !== 200 &&
    status !== 201
  ) {

    throw new Error(
      "GitHub upload failed (" +
      status +
      "): " +
      responseText
    );

  }


  return {

    success: true,

    fileName:
      finalFileName

  };

}