const XLSX = require("xlsx");

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function findEmailKeyFromHeaders(row) {
  const keys = Object.keys(row);
  return keys.find((k) => k.trim().toLowerCase() === "contact_email");
}

function validateEmail(email) {
  return typeof email === "string" && emailRegex.test(email.trim());
}

function validateRows(rows) {
  if (!rows || rows.length === 0) return [];

  const emailKey = findEmailKeyFromHeaders(rows[0]);

  const groupedBySchool = {};

  rows.forEach((row, index) => {
    let newRow = { ...row };
    const statusMessages = [];
    let errorCount = 0;

    let schoolName = row["SCHOOL_NAME"]?.toString().trim();
    const normalizedSchoolName = schoolName?.toLowerCase();

    const orgCode = row["ORG_CODE"]?.toString().trim();
    if (orgCode) {
      const normalizedOrgCode = orgCode.toUpperCase().replace(/-/g, "");
      if (normalizedOrgCode !== `U-${normalizedOrgCode}`) {
        statusMessages.push(
          `ORG_CODE should be in format U-XXXXX got ${orgCode}`
        );
        errorCount++;
      }
    }

    const orgPin = row["ORG_PIN"]?.toString().trim();
    if (orgPin) {
      if (orgPin && !/^\d{6}$/.test(orgPin)) {
        statusMessages.push("ORG_PIN should be a 6-digit numeric value");
        errorCount++;
      }
    }

    const admissionYear = row["ADMISSION_YEAR"]?.toString().trim();
    if (admissionYear && !/^\d{4}$/.test(admissionYear)) {
      statusMessages.push("ADMISSION_YEAR should be a 4-digit numeric value");
      errorCount++;
    }

    const abcAccountId = row["ABC_ACCOUNT_ID"]?.toString().trim();
    if (abcAccountId && !/^\d{12}$/.test(abcAccountId)) {
      statusMessages.push("ABC_ACCOUNT_ID should be a 12-digit numeric value");
      errorCount++;
    }

    const dob = row["DOB"]?.toString().trim();
    if (dob && !/^\d{2}\/\d{2}\/\d{4}$/.test(dob)) {
      statusMessages.push("DOB should be in format DD/MM/YYYY");
      errorCount++;
    }

    const gender = row["GENDER"]?.toString().trim();
    if (gender && !/^[MFTX]$/i.test(gender)) {
      statusMessages.push("GENDER should be M/F/T/X");
      errorCount++;
    }

    const caste = row["CASTE"]?.toString().trim().toUpperCase();
    if (caste && !["GEN", "OBC", "OBCNC", "SC", "ST"].includes(caste)) {
      statusMessages.push("CASTE should be GEN/OBC/OBCNC/SC/ST");
      errorCount++;
    }

    const religion = row["RELIGION"]?.toString().trim().toUpperCase();
    if (
      religion &&
      !["H", "I", "C", "S", "B", "J", "Z", "O"].includes(religion)
    ) {
      statusMessages.push("RELIGION should be H/I/C/S/B/J/Z/O");
      errorCount++;
    }

    const nationality = row["NATIONALITY"]?.toString().trim().toUpperCase();
    if (nationality && nationality !== "IN") {
      statusMessages.push("NATIONALITY should be IN for India");
      errorCount++;
    }

    const disabilityStatus = row["DISABILITY_STATUS"]
      ?.toString()
      .trim()
      .toUpperCase();
    if (
      disabilityStatus &&
      disabilityStatus !== "Y" &&
      disabilityStatus !== "N"
    ) {
      statusMessages.push("DISABILITY_STATUS should be Y or N");
      errorCount++;
    }

    const result = row["RESULT"]?.toString().trim().toUpperCase();
    if (
      result &&
      !["PASS", "FAIL", "QUALIFIED", "COMPULSORY REPEAT", "ABSENT"].includes(
        result
      )
    ) {
      statusMessages.push(
        "RESULT should be PASS/FAIL/QUALIFIED/COMPULSORY REPEAT/ABSENT"
      );
      errorCount++;
    }

    const resultTh = row["RESULT_TH"]?.toString().trim().toUpperCase();
    if (
      resultTh &&
      !["PASS", "FAIL", "QUALIFIED", "COMPULSORY REPEAT", "ABSENT"].includes(
        resultTh
      )
    ) {
      statusMessages.push(
        "RESULT_TH should be PASS/FAIL/QUALIFIED/COMPULSORY REPEAT/ABSENT"
      );
      errorCount++;
    }

    const marksRecStatus = row["MRKS_REC_STATUS"]
      ?.toString()
      .trim()
      .toUpperCase();
    if (marksRecStatus && !["O", "M", "C"].includes(marksRecStatus)) {
      statusMessages.push("MRKS_REC_STATUS should be O/M/C");
      errorCount++;
    }

    const resultPr = row["RESULT_PR"]?.toString().trim().toUpperCase();
    if (
      resultPr &&
      !["PASS", "FAIL", "QUALIFIED", "COMPULSORY REPEAT", "ABSENT"].includes(
        resultPr
      )
    ) {
      statusMessages.push(
        "RESULT_PR should be PASS/FAIL/QUALIFIED/COMPULSORY REPEAT/ABSENT"
      );
      errorCount++;
    }

    const yearOfExam = row["YEAR"]?.toString().trim();
    if (yearOfExam && !/^\d{4}$/.test(yearOfExam)) {
      statusMessages.push("YEAR should be a valid year in YYYY format");
      errorCount++;
    }

    const percentage = row["PERCENT"]?.toString().trim();
    if (percentage && !/^[\d.]+$/.test(percentage)) {
      statusMessages.push("PERCENT should be a valid number");
      errorCount++;
    }

    const dorCertificateDate = row["DOR"]?.toString().trim();
    if (
      dorCertificateDate &&
      !/^\d{2}\/\d{2}\/\d{4}$/.test(dorCertificateDate)
    ) {
      statusMessages.push("DOR should be in DD/MM/YYYY format");
      errorCount++;
    }

    const dorCertificateIssueDate = row["DOI"]?.toString().trim();
    if (
      dorCertificateIssueDate &&
      !/^\d{2}\/\d{2}\/\d{4}$/.test(dorCertificateIssueDate)
    ) {
      statusMessages.push("DOI should be in DD/MM/YYYY format");
      errorCount++;
    }

    const dovCertificateValidityDate = row["DOV"]?.toString().trim();
    if (
      dovCertificateValidityDate &&
      !/^\d{2}\/\d{2}\/\d{4}$/.test(dovCertificateValidityDate)
    ) {
      statusMessages.push("DOV should be in DD/MM/YYYY format");
      errorCount++;
    }

    const doeDate = row["DOE"]?.toString().trim();
    if (doeDate && !/^\d{2}\/\d{2}\/\d{4}$/.test(doeDate)) {
      statusMessages.push("DOE should be in DD/MM/YYYY format");
      errorCount++;
    }

    const dopDate = row["DOP"]?.toString().trim();
    if (dopDate && !/^\d{2}\/\d{2}\/\d{4}$/.test(dopDate)) {
      statusMessages.push("DOP should be in DD/MM/YYYY format");
      errorCount++;
    }

    const doqDate = row["DOQ"]?.toString().trim();
    if (doqDate && !/^\d{2}\/\d{2}\/\d{4}$/.test(doqDate)) {
      statusMessages.push("DOQ should be in DD/MM/YYYY format");
      errorCount++;
    }

    const dosDate = row["DOS"]?.toString().trim();
    if (dosDate && !/^\d{2}\/\d{2}\/\d{4}$/.test(dosDate)) {
      statusMessages.push("DOS should be in DD/MM/YYYY format");
      errorCount++;
    }

    const fieldsToValidate = [
      { key: "TOT", message: "TOT should be a valid number" },
      { key: "TOT_MIN", message: "TOT_MIN should be a valid number" },
      { key: "TOT_MRKS", message: "TOT_MRKS should be a valid number" },
      { key: "TOT_MRKS_MIN", message: "TOT_MRKS_MIN should be a valid number" },
      { key: "TOT_TH_MAX", message: "TOT_TH_MAX should be a valid number" },
      { key: "TOT_TH_MIN", message: "TOT_TH_MIN should be a valid number" },
      { key: "TOT_TH_MRKS", message: "TOT_TH_MRKS should be a valid number" },
      { key: "TOT_PR_MAX", message: "TOT_PR_MAX should be a valid number" },
      { key: "TOT_PR_MIN", message: "TOT_PR_MIN should be a valid number" },
      { key: "TOT_PR_MRKS", message: "TOT_PR_MRKS should be a valid number" },
      { key: "TOT_CE_MAX", message: "TOT_CE_MAX should be a valid number" },
      { key: "TOT_CE_MIN", message: "TOT_CE_MIN should be a valid number" },
      { key: "TOT_CE_MRKS", message: "TOT_CE_MRKS should be a valid number" },
      { key: "TOT_VV_MAX", message: "TOT_VV_MAX should be a valid number" },
      { key: "TOT_VV_MIN", message: "TOT_VV_MIN should be a valid number" },
      { key: "TOT_VV_MRKS", message: "TOT_VV_MRKS should be a valid number" },
      {
        key: "TOT_PR_CE_MAX",
        message: "TOT_PR_CE_MAX should be a valid number",
      },
      {
        key: "TOT_PR_CE_MRKS",
        message: "TOT_PR_CE_MRKS should be a valid number",
      },
      {
        key: "TOT_TH_CE_MAX",
        message: "TOT_TH_CE_MAX should be a valid number",
      },
      {
        key: "TOT_TH_CE_MRKS",
        message: "TOT_TH_CE_MRKS should be a valid number",
      },
      {
        key: "PREV_TOT_MRKS",
        message: "PREV_TOT_MRKS should be a valid number",
      },
      {
        key: "PREV_TOT_MRKS_MAX",
        message: "PREV_TOT_MRKS_MAX should be a valid number",
      },
      {
        key: "PREV_TOT_MRKS_MIN",
        message: "PREV_TOT_MRKS_MIN should be a valid number",
      },
      {
        key: "GRAND_TOT_MAX",
        message: "GRAND_TOT_MAX should be a valid number",
      },
      {
        key: "GRAND_TOT_MIN",
        message: "GRAND_TOT_MIN should be a valid number",
      },
      {
        key: "GRAND_TOT_MRKS",
        message: "GRAND_TOT_MRKS should be a valid number",
      },
      {
        key: "GRAND_TOT_GRADE_POINTS",
        message: "GRAND_TOT_GRADE_POINTS should be a valid number",
      },
      {
        key: "GRAND_TOT_CREDIT",
        message: "GRAND_TOT_CREDIT should be a valid number",
      },
      {
        key: "GRAND_TOT_CREDIT_POINTS",
        message: "GRAND_TOT_CREDIT_POINTS should be a valid number",
      },
      {
        key: "OT_CGPA_MINOR",
        message: "OT_CGPA_MINOR should be a valid number",
      },
      {
        key: "TOT_CREDIT_MINOR",
        message: "TOT_CREDIT_MINOR should be a valid number",
      },
      {
        key: "FINAL_GMAX_TOTAL",
        message: "FINAL_GMAX_TOTAL should be a valid number",
      },
      {
        key: "CGPA_SCALE",
        message: "CGPA_SCALE should be a valid number",
      },
    ];

    fieldsToValidate.forEach(({ key, message }) => {
      const value = row[key]?.toString().trim();
      if (value && !/^\d+$/.test(value)) {
        statusMessages.push(message);
        errorCount++;
      }
    });

    const fieldsValidation = [
      {
        key: "CGPA",
        pattern: /^[a-zA-Z0-9/.\/]+$/i,
        message:
          "CGPA should allow letters, numbers, periods and forward slashes",
      },
      {
        key: "GPA",
        pattern: /^[a-zA-Z0-9/.\/]+$/i,
        message:
          "GPA should allow letters, numbers, periods and forward slashes",
      },
      {
        key: "SGPA",
        pattern: /^[a-zA-Z0-9/.\/]+$/i,
        message:
          "SGPA should allow letters, numbers, periods and forward slashes",
      },
      {
        key: "INTR_CGPA_FIRST",
        pattern: /^[a-zA-Z0-9/.\/]+$/i,
        message:
          "INTR_CGPA_FIRST should allow letters, numbers, periods and forward slashes",
      },
      {
        key: "INTR_CGPA_SECOND",
        pattern: /^[a-zA-Z0-9/.\/]+$/i,
        message:
          "INTR_CGPA_SECOND should allow letters, numbers, periods and forward slashes",
      },
      {
        key: "SUB1MAX",
        pattern: /^\d+$/,
        message: "SUB1MAX should be a numeric value",
      },
      {
        key: "SUB1MIN",
        pattern: /^\d+$/,
        message: "SUB1MIN should be a numeric value",
      },
      {
        key: "SUB1_TH_MAX",
        pattern: /^\d+$/,
        message: "SUB1_TH_MAX should be a numeric value",
      },
      {
        key: "SUB1_TH_MIN",
        pattern: /^\d+$/,
        message: "SUB1_TH_MIN should be a numeric value",
      },
      {
        key: "SUB1_PR_MAX",
        pattern: /^\d+$/,
        message: "SUB1_PR_MAX should be a numeric value",
      },
      {
        key: "SUB1_PR_MIN",
        pattern: /^\d+$/,
        message: "SUB1_PR_MIN should be a numeric value",
      },
      {
        key: "SUB1_CE_MAX",
        pattern: /^\d+$/,
        message: "SUB1_CE_MAX should be a numeric value",
      },
      {
        key: "SUB1_CE_MIN",
        pattern: /^\d+$/,
        message: "SUB1_CE_MIN should be a numeric value",
      },
      {
        key: "SUB1_VV_MAX",
        pattern: /^\d+$/,
        message: "SUB1_VV_MAX should be a numeric value",
      },
      {
        key: "SUB1_VV_MIN",
        pattern: /^\d+$/,
        message: "SUB1_VV_MIN should be a numeric value",
      },
      {
        key: "SUB1_CE_WEIGHT_MRKS",
        pattern: /^\d+(\.\d+)?$/,
        message:
          "SUB1_CE_WEIGHT_MRKS should be a numeric value (decimal allowed)",
      },
      {
        key: "SUB1_CE_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_CE_MRKS should be a numeric value",
      },
      {
        key: "SUB1_CE1_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_CE1_MRKS should be a numeric value",
      },
      {
        key: "SUB1_CE2_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_CE2_MRKS should be a numeric value",
      },
      {
        key: "SUB1_CE3_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_CE3_MRKS should be a numeric value",
      },
      {
        key: "SUB1_CE4_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_CE4_MRKS should be a numeric value",
      },
      {
        key: "SUB1_VV_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_VV_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER4_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER4_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_PR_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_PR_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_PR_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_PR_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_PR_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_PR_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_CE_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_CE_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_CE_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_CE_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_CE_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_CE_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_CE_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_CE_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER1_PR_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER1_PR_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_CE_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_CE_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER2_PR_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER2_PR_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_CE_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_CE_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_PAPER3_PR_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_PAPER3_PR_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MAX_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MAX_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MAX_CE_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MAX_CE_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MAX_PR_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MAX_PR_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MIN_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MIN_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MIN_CE_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MIN_CE_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_MIN_PR_MRKS_SH",
        pattern: /^\d+$/,
        message: "SUB1_MIN_PR_MRKS_SH should be a numeric value",
      },
      {
        key: "SUB1_LAB1_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_LAB1_MRKS should be a numeric value",
      },
      {
        key: "SUB1_LAB2_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_LAB2_MRKS should be a numeric value",
      },
      {
        key: "SUB1_LAB3_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_LAB3_MRKS should be a numeric value",
      },
      {
        key: "SUB1_LAB4_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_LAB4_MRKS should be a numeric value",
      },
      {
        key: "SUB1_REPORT_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_REPORT_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PRO_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PRO_MRKS should be a numeric value",
      },
      {
        key: "SUB1_PRO_CE_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_PRO_CE_MRKS should be a numeric value",
      },
      {
        key: "SUB1_TEE_PR_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_TEE_PR_MRKS should be a numeric value",
      },
      {
        key: "SUB1_TEE_TH_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_TEE_TH_MRKS should be a numeric value",
      },
      {
        key: "SUB1_TEE_PR_GRADE",
        pattern: /^\d+$/,
        message: "SUB1_TEE_PR_GRADE should be a numeric value",
      },
      {
        key: "SUB1_TEE_TH_GRADE",
        pattern: /^\d+$/,
        message: "SUB1_TEE_TH_GRADE should be a numeric value",
      },
      {
        key: "SUB1_TEE_WEIGHT_MRKS",
        pattern: /^\d+$/,
        message: "SUB1_TEE_WEIGHT_MRKS should be a numeric value",
      },
      {
        key: "SUB1_TYPE",
        pattern: /^(M|E|A|B)$/,
        message: "SUB1_TYPE should be M, E, A or B",
      },
      {
        key: "SUB1_TOT",
        pattern: /^\d+$/,
        message: "SUB1_TOT should be a numeric value",
      },
      {
        key: "SUB1_CE_TOT",
        pattern: /^\d+$/,
        message: "SUB1_CE_TOT should be a numeric value",
      },
      {
        key: "SUB1_PR_TOT",
        pattern: /^\d+$/,
        message: "SUB1_PR_TOT should be a numeric value",
      },
      {
        key: "SUB1_STATUS",
        pattern: /^(Pass|Fail)$/,
        message: "SUB1_STATUS should be Pass or Fail",
      },
      {
        key: "SUB1_PAPER1_STATUS",
        pattern: /^(Pass|Fail|Reappear)$/,
        message: "SUB1_PAPER1_STATUS should be Pass or Fail or Reappear",
      },
      {
        key: "SUB1_PAPER2_STATUS",
        pattern: /^(Pass|Fail|Reappear)$/,
        message: "SUB1_PAPER2_STATUS should be Pass or Fail or Reappear",
      },
      {
        key: "SUB1_PAPER3_STATUS",
        pattern: /^(Pass|Fail|Reappear)$/,
        message: "SUB1_PAPER3_STATUS should be Pass or Fail or Reappear",
      },
      {
        key: "SUB1_PAPER4_STATUS",
        pattern: /^(Pass|Fail|Reappear)$/,
        message: "SUB1_PAPER4_STATUS should be Pass or Fail or Reappear",
      },
    ];

    fieldsValidation.forEach(({ key, pattern, message }) => {
      const value = row[key]?.toString().trim();
      if (value && !pattern.test(value)) {
        statusMessages.push(message);
        errorCount++;
      }
    });

    if (emailKey) {
      const email = row[emailKey]?.toString().trim();
      if (email && !validateEmail(email)) {
        statusMessages.push("Invalid Email");
        errorCount++;
      }
    }

    newRow = {
      Status: statusMessages.length ? statusMessages.join("\n") : "Valid",
      ErrorsCount: errorCount,
      ...newRow,
    };

    if (!normalizedSchoolName) {
      if (!groupedBySchool["__UNKNOWN__"])
        groupedBySchool["__UNKNOWN__"] = {
          schoolDisplayName: "Unknown",
          rows: [],
        };
      groupedBySchool["__UNKNOWN__"].rows.push(newRow);
    } else {
      if (!groupedBySchool[normalizedSchoolName]) {
        groupedBySchool[normalizedSchoolName] = {
          schoolDisplayName: schoolName,
          rows: [],
        };
      }
      groupedBySchool[normalizedSchoolName].rows.push(newRow);
    }
  });

  return groupedBySchool;
}

function parseExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const firstRow = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    range: 0,
    header: 1,
  })[0];

  const jsonWithHeaderRow = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    range: 1,
  });

  const groupedData = validateRows(jsonWithHeaderRow);
  return { firstRow, groupedData };
}

module.exports = { parseExcel, validateRows };
