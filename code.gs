/* =========================================================
   MEDICINE HUB
   ========================================================= */

const MEDICINE_SPREADSHEET_ID =
  "1Lpm8iW4RNnkvw4MiqpzfGi-H5Uh9dTIRztARmE4ZpJg";

const MEDICINE_SHEET_NAME = "Medicines";
const SETTINGS_SHEET_NAME = "Settings";

const DELETE_PASSCODE = "12345";


/* =========================================================
   WEB APP
========================================================= */

function doGet() {

  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Medicine Hub")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* =========================================================
   INCLUDE
========================================================= */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* =========================================================
   MEDICINES SHEET
========================================================= */

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


/* =========================================================
   GET MEDICINES
========================================================= */

function getMedicines() {

  const sheet =
    getMedicineSheet();

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        5
      )
      .getValues();


  return values

    .filter(function(row) {

      return String(
        row[0] || ""
      ).trim() !== "";

    })

    .map(function(row) {

      return {

        name:
          String(
            row[0] || ""
          ).trim(),

        category:
          String(
            row[1] || ""
          ).trim(),

        onlineLink:
          String(
            row[2] || ""
          ).trim(),

        genericName:
          String(
            row[3] || ""
          ).trim(),

        favorite:
          isFavoriteValue(
            row[4]
          )

      };

    });

}


/* =========================================================
   FAVORITE VALUE
========================================================= */

function isFavoriteValue(value) {

  if (value === true) {

    return true;

  }


  if (value === 1) {

    return true;

  }


  const text =
    String(
      value || ""
    )
    .trim()
    .toLowerCase();


  return (
    text === "true" ||
    text === "yes" ||
    text === "y" ||
    text === "1" ||
    text === "star" ||
    text === "⭐"
  );

}


/* =========================================================
   SETTINGS SHEET
========================================================= */

function getSettingsSheet() {

  const ss =
    SpreadsheetApp.openById(
      MEDICINE_SPREADSHEET_ID
    );


  let sheet =
    ss.getSheetByName(
      SETTINGS_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        SETTINGS_SHEET_NAME
      );

  }


  /*
   * Make sure headers exist.
   */

  if (
    sheet.getRange("A1").getValue() !==
    "Category"
  ) {

    sheet
      .getRange("A1")
      .setValue("Category");

  }


  if (
    sheet.getRange("B1").getValue() !==
    "Icon"
  ) {

    sheet
      .getRange("B1")
      .setValue("Icon");

  }


  return sheet;

}


/* =========================================================
   GET CATEGORIES
========================================================= */

function getCategories() {

  const sheet =
    getSettingsSheet();

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  /*
   * Read Category + Icon.
   */

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  return values

    .filter(function(row) {

      return String(
        row[0] || ""
      ).trim() !== "";

    })

    .map(function(row) {

      return {

        name:
          String(
            row[0] || ""
          ).trim(),

        icon:
          String(
            row[1] || ""
          ).trim() || "💊"

      };

    });

}


/* =========================================================
   ADD MEDICINE
========================================================= */

function addMedicine(medicine) {

  if (!medicine) {

    throw new Error(
      "Medicine data is missing."
    );

  }


  const name =
    String(
      medicine.name || ""
    ).trim();


  if (!name) {

    throw new Error(
      "Medicine name is required."
    );

  }


  const sheet =
    getMedicineSheet();


  /*
   * DUPLICATE CHECK
   */

  const duplicate =
    findMedicineByName(
      sheet,
      name
    );


  if (duplicate) {

    return {

      success: false,

      duplicate: true,

      message:
        'Medicine "' +
        name +
        '" already exists.'

    };

  }


  sheet.appendRow([

    name,

    String(
      medicine.category || ""
    ).trim(),

    String(
      medicine.onlineLink || ""
    ).trim(),

    String(
      medicine.genericName || ""
    ).trim(),

    medicine.favorite === true

  ]);


  return {

    success: true

  };

}


/* =========================================================
   FIND MEDICINE BY NAME
========================================================= */

function findMedicineByName(
  sheet,
  name,
  excludeName
) {

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return null;

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        5
      )
      .getValues();


  const target =
    String(
      name || ""
    )
    .trim()
    .toLowerCase();


  const excluded =
    String(
      excludeName || ""
    )
    .trim()
    .toLowerCase();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const existing =
      String(
        values[i][0] || ""
      )
      .trim();


    if (!existing) {

      continue;

    }


    const existingLower =
      existing.toLowerCase();


    if (
      existingLower === target &&
      existingLower !== excluded
    ) {

      return {

        row:
          i + 2,

        name:
          existing

      };

    }

  }


  return null;

}


/* =========================================================
   FIND MEDICINE ROW
========================================================= */

function findMedicineRowByName(
  sheet,
  name
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


  const target =
    String(
      name || ""
    )
    .trim()
    .toLowerCase();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const current =
      String(
        values[i][0] || ""
      )
      .trim()
      .toLowerCase();


    if (
      current === target
    ) {

      return i + 2;

    }

  }


  return -1;

}


/* =========================================================
   UPDATE MEDICINE
========================================================= */

function updateMedicine(
  medicine,
  originalName
) {

  if (!medicine) {

    throw new Error(
      "Medicine data is missing."
    );

  }


  const name =
    String(
      medicine.name || ""
    ).trim();


  if (!name) {

    throw new Error(
      "Medicine name is required."
    );

  }


  const sheet =
    getMedicineSheet();


  /*
   * Duplicate check when changing name.
   */

  const duplicate =
    findMedicineByName(
      sheet,
      name,
      originalName
    );


  if (duplicate) {

    return {

      success: false,

      duplicate: true,

      message:
        'Medicine "' +
        name +
        '" already exists.'

    };

  }


  const rowNumber =
    findMedicineRowByName(
      sheet,
      originalName || name
    );


  if (rowNumber === -1) {

    throw new Error(
      "Medicine not found."
    );

  }


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      5
    )
    .setValues([

      [

        name,

        String(
          medicine.category || ""
        ).trim(),

        String(
          medicine.onlineLink || ""
        ).trim(),

        String(
          medicine.genericName || ""
        ).trim(),

        medicine.favorite === true

      ]

    ]);


  return {

    success: true

  };

}


/* =========================================================
   TOGGLE FAVORITE
========================================================= */

function toggleMedicineFavorite(
  name,
  favorite
) {

  const sheet =
    getMedicineSheet();


  const rowNumber =
    findMedicineRowByName(
      sheet,
      name
    );


  if (rowNumber === -1) {

    throw new Error(
      "Medicine not found."
    );

  }


  /*
   * Explicitly write TRUE/FALSE
   * to column E.
   */

  sheet
    .getRange(
      rowNumber,
      5
    )
    .setValue(
      favorite === true
    );


  SpreadsheetApp.flush();


  return {

    success: true,

    favorite:
      favorite === true

  };

}


/* =========================================================
   DELETE MEDICINE
========================================================= */

function deleteMedicine(
  name,
  passcode
) {

  if (
    String(passcode) !==
    DELETE_PASSCODE
  ) {

    throw new Error(
      "Incorrect passcode."
    );

  }


  const sheet =
    getMedicineSheet();


  const rowNumber =
    findMedicineRowByName(
      sheet,
      name
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


/* =========================================================
   ADD CATEGORY
========================================================= */

function addCategory(
  category,
  icon
) {

  const name =
    String(
      category || ""
    ).trim();


  const categoryIcon =
    String(
      icon || ""
    ).trim() || "💊";


  if (!name) {

    throw new Error(
      "Category name is required."
    );

  }


  const sheet =
    getSettingsSheet();


  const categories =
    getCategories();


  const duplicate =
    categories.some(
      function(item) {

        return (
          item.name
            .toLowerCase() ===
          name.toLowerCase()
        );

      }
    );


  if (duplicate) {

    return {

      success: false,

      duplicate: true,

      message:
        'Category "' +
        name +
        '" already exists.'

    };

  }


  sheet.appendRow([

    name,

    categoryIcon

  ]);


  return {

    success: true

  };

}


/* =========================================================
   UPDATE CATEGORY
========================================================= */

function updateCategory(
  originalName,
  newName,
  icon
) {

  const oldName =
    String(
      originalName || ""
    ).trim();


  const name =
    String(
      newName || ""
    ).trim();


  const categoryIcon =
    String(
      icon || ""
    ).trim() || "💊";


  if (!oldName) {

    throw new Error(
      "Original category is missing."
    );

  }


  if (!name) {

    throw new Error(
      "Category name is required."
    );

  }


  const sheet =
    getSettingsSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "Category not found."
    );

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  /*
   * Check duplicate name.
   */

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const existing =
      String(
        values[i][0] || ""
      )
      .trim();


    if (
      existing.toLowerCase() ===
      name.toLowerCase() &&
      existing.toLowerCase() !==
      oldName.toLowerCase()
    ) {

      return {

        success: false,

        duplicate: true,

        message:
          'Category "' +
          name +
          '" already exists.'

      };

    }

  }


  /*
   * Find original category.
   */

  let rowNumber = -1;


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const existing =
      String(
        values[i][0] || ""
      )
      .trim()
      .toLowerCase();


    if (
      existing ===
      oldName.toLowerCase()
    ) {

      rowNumber =
        i + 2;

      break;

    }

  }


  if (rowNumber === -1) {

    throw new Error(
      "Category not found."
    );

  }


  /*
   * Update Settings.
   */

  sheet
    .getRange(
      rowNumber,
      1,
      1,
      2
    )
    .setValues([

      [
        name,
        categoryIcon
      ]

    ]);


  /*
   * If category name changed,
   * update medicines using the
   * old category.
   */

  updateMedicineCategories(
    oldName,
    name
  );


  return {

    success: true

  };

}


/* =========================================================
   UPDATE MEDICINES AFTER CATEGORY RENAME
========================================================= */

function updateMedicineCategories(
  oldName,
  newName
) {

  const sheet =
    getMedicineSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return;

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        5
      )
      .getValues();


  let changed = false;


  values.forEach(
    function(row) {

      const category =
        String(
          row[1] || ""
        ).trim();


      if (
        category.toLowerCase() ===
        oldName.toLowerCase()
      ) {

        row[1] =
          newName;

        changed = true;

      }

    }
  );


  if (changed) {

    sheet
      .getRange(
        2,
        1,
        values.length,
        5
      )
      .setValues(
        values
      );

  }

}


/* =========================================================
   DELETE CATEGORY
========================================================= */

function deleteCategory(
  category
) {

  const name =
    String(
      category || ""
    ).trim();


  const medicineSheet =
    getMedicineSheet();


  const medicineLastRow =
    medicineSheet.getLastRow();


  /*
   * Do not delete a category
   * that is currently being used.
   */

  if (
    medicineLastRow >= 2
  ) {

    const medicines =
      medicineSheet
        .getRange(
          2,
          1,
          medicineLastRow - 1,
          5
        )
        .getValues();


    const used =
      medicines.some(
        function(row) {

          return (
            String(
              row[1] || ""
            )
            .trim()
            .toLowerCase() ===
            name.toLowerCase()
          );

        }
      );


    if (used) {

      throw new Error(
        'Cannot delete "' +
        name +
        '" because medicines are using this category.'
      );

    }

  }


  const sheet =
    getSettingsSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "Category not found."
    );

  }


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

    const existing =
      String(
        values[i][0] || ""
      )
      .trim()
      .toLowerCase();


    if (
      existing ===
      name.toLowerCase()
    ) {

      sheet.deleteRow(
        i + 2
      );


      return {

        success: true

      };

    }

  }


  throw new Error(
    "Category not found."
  );

}