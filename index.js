const xlsx = require("xlsx");
const fs = require("fs");
require("dotenv").config();
const contentfulImport = require("contentful-import");

// contentful configs
const options = {
  contentFile: "./contentful-import.json",
  spaceId: process.env.SPACE_ID,
  managementToken: process.env.MANAGEMENT_TOKEN,
  environmentId: process.env.ENVIRONMENT_ID,
};

console.log(options);

// Load the Excel file
const workbook = xlsx.readFile("Feedback.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the worksheet to JSON
const rawJson = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

// Extract the headers from the second row
const fieldMapping = rawJson[1];

// Process the data starting from the third row
const mappedJson = rawJson.slice(2).map((row) => {
  let mappedRow = {};
  Object.entries(fieldMapping).forEach(([key, newKey]) => {
    mappedRow[newKey] = row[key];
  });
  mappedRow["Center Name"] =
    row["The Face Flex Cosmetic and Personal Care Limited"];
  return mappedRow;
});

//contentful mapping object
const contentfulMapping = {
  "Guest Name": "customerName",
  Rating: "rating",
  Tags: "title",
  "Guest Comments": "description",
  Service: "lastWorkout",
};

//creating the final json required for contentful
const structuredJson = mappedJson.map((row) => {
  let mappedRow = {};
  for (let [excelField, jsonField] of Object.entries(contentfulMapping)) {
    mappedRow[jsonField] = row[excelField];
  }
  return mappedRow;
});

fs.writeFileSync(
  "csv-to-flexer.json",
  JSON.stringify(structuredJson, null, 2),
  "utf-8"
);

const finalJson = structuredJson.map((entry) => {
  return {
    metadata: {
      tags: [],
    },
    sys: {
      space: {
        sys: {
          type: "Link",
          linkType: "Space",
          id: options.spaceId,
        },
      },
      id: "",
      type: "Entry",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      environment: {
        sys: {
          id: "master",
          type: "Link",
          linkType: "Environment",
        },
      },
      revision: 1,
      contentType: {
        sys: {
          type: "Link",
          linkType: "ContentType",
          id: "flexerCard",
        },
      },
    },
    fields: {
      rating: {
        "en-US": entry.rating,
      },
      title: {
        "en-US": entry.title,
      },
      description: {
        "en-US": entry.description,
      },
      customerName: {
        "en-US": entry.customerName,
      },
      lastWorkout: {
        "en-US": entry.lastWorkout,
      },
    },
  };
});

const importJson = {
  entries: finalJson,
};

// Write the mapped JSON to a file
fs.writeFileSync(
  "contentful-import.json",
  JSON.stringify(importJson, null, 2),
  "utf-8"
);

console.log("JSON file has been created successfully.");

contentfulImport(options)
  .then(() => {
    console.log("Data imported successfully");
  })
  .catch((err) => {
    console.log("Oh no! Some errors occurred!", err);
  });
