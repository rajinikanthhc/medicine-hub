const SHEET_NAME = "Medicines";

const CATEGORY_LIST = [
  "Fever",
  "Cough & Cold",
  "Pain Relief",
  "Allergy",
  "Gastric & Acidity",
  "Antibiotics",
  "Vitamins & Supplements",
  "Diabetes",
  "Blood Pressure",
  "Skin",
  "Eye & Ear",
  "Digestive",
  "Other"
];


/* ==================================================
   WEB APP
================================================== */

function doGet() {

  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Medicine Hub")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}


/* ==================================================
   INCLUDE HTML FILES
================================================== */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* ==================================================
   GET MEDICINES
================================================== */

function getMedicines() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Medicines sheet not found.");
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  return data.slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({

      id: Number(row[0]),

      name: row[1] || "",

      genericName: row[2] || "",

      usedFor: row[3] || "",

      medicineUse: row[4] || "",

      description: row[5] || "",

      photo: row[6] || "",

      favorite:
        String(row[7]).toLowerCase() === "true" ||
        String(row[7]) === "1" ||
        String(row[7]).toLowerCase() === "yes",

      category: row[8] || ""

    }));

}


/* ==================================================
   CATEGORY LIST
================================================== */

function getCategories() {

  return CATEGORY_LIST;

}


/* ==================================================
   ADD / EDIT MEDICINE
================================================== */

function saveMedicine(medicine) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Medicines sheet not found.");
  }


  const data = sheet.getDataRange().getValues();


  /* -----------------------------------------------
     CHECK DUPLICATE NAME
  ------------------------------------------------ */

  const enteredName =
    String(medicine.name || "")
      .trim()
      .toLowerCase();


  if (!enteredName) {
    throw new Error("Medicine name is required.");
  }


  for (let i = 1; i < data.length; i++) {

    const existingId =
      Number(data[i][0]);

    const existingName =
      String(data[i][1] || "")
        .trim()
        .toLowerCase();


    if (
      existingName === enteredName &&
      existingId !== Number(medicine.id)
    ) {

      return {
        success: false,
        duplicate: true,
        message:
          "Medicine already exists: " +
          data[i][1]
      };

    }

  }


  /* -----------------------------------------------
     EDIT
  ------------------------------------------------ */

  if (medicine.id) {

    for (let i = 1; i < data.length; i++) {

      if (
        Number(data[i][0]) ===
        Number(medicine.id)
      ) {

        sheet.getRange(i + 1, 1, 1, 9)
          .setValues([[
            Number(medicine.id),
            medicine.name || "",
            medicine.genericName || "",
            medicine.usedFor || "",
            medicine.medicineUse || "",
            medicine.description || "",
            medicine.photo || "",
            medicine.favorite ? true : false,
            medicine.category || ""
          ]]);

        return {
          success: true,
          message: "Medicine updated successfully."
        };

      }

    }

    throw new Error("Medicine not found.");

  }


  /* -----------------------------------------------
     CREATE NEW ID
     LOWEST AVAILABLE NUMBER
  ------------------------------------------------ */

  const usedIds = data
    .slice(1)
    .map(row => Number(row[0]))
    .filter(id => !isNaN(id) && id > 0);


  let newId = 1;

  while (usedIds.includes(newId)) {
    newId++;
  }


  sheet.appendRow([

    newId,

    medicine.name || "",

    medicine.genericName || "",

    medicine.usedFor || "",

    medicine.medicineUse || "",

    medicine.description || "",

    medicine.photo || "",

    medicine.favorite ? true : false,

    medicine.category || ""

  ]);


  return {

    success: true,

    message:
      "Medicine added successfully.",

    id: newId

  };

}


/* ==================================================
   DELETE MEDICINE
================================================== */

function deleteMedicine(id) {

  const PASSWORD = "12345";


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  if (!sheet) {
    throw new Error("Medicines sheet not found.");
  }


  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    if (
      Number(data[i][0]) ===
      Number(id)
    ) {

      sheet.deleteRow(i + 1);

      return {
        success: true,
        message: "Medicine deleted."
      };

    }

  }


  throw new Error("Medicine not found.");

}


/* ==================================================
   TOGGLE FAVORITE
================================================== */

function toggleFavorite(id) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);


  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    if (
      Number(data[i][0]) ===
      Number(id)
    ) {

      const current =
        String(data[i][7]).toLowerCase() === "true" ||
        String(data[i][7]) === "1" ||
        String(data[i][7]).toLowerCase() === "yes";


      sheet
        .getRange(i + 1, 8)
        .setValue(!current);


      return {
        success: true,
        favorite: !current
      };

    }

  }


  throw new Error("Medicine not found.");

}