const GITHUB_RAW_BASE = "https://raw.githubusercontent.com/giantswarm/feedback-templates/refs/heads/main/";

const ROLES_MANIFEST_FILE = "roles.json";

const TEMPLATES_FOLDER = "questions/";

const PROPERTY_PREFIX = "EmployeeFeedback.";

const FORM_ID_KEY = PROPERTY_PREFIX + "formId";

const FOLDER_ID_KEY = PROPERTY_PREFIX + "folderId";

const RESPONSES_SHEET_ID_KEY = PROPERTY_PREFIX + "responsesSheetId";

/**
 * Convenience function to configure script properties.
 *
 * USAGE:
 *   clasp run 'setProperties' --params '[{"EmployeeFeedback.formId": "YOUR_FORM_ID"}, false]'
 */
function setProperties(properties, deleteAllOthers) {
  TriggerUtil.setProperties(properties, deleteAllOthers);
}

/**
 * Install the onFormSubmit trigger.
 * Requires EmployeeFeedback.formId to be set in Script Properties first.
 *
 * USAGE:
 *   clasp run 'install'
 */
function install() {
  const formId = PropertiesService.getScriptProperties().getProperty(FORM_ID_KEY);
  if (!formId) throw new Error("Set " + FORM_ID_KEY + " in Script Properties before running install()");

  uninstall();

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(formId)
    .onFormSubmit()
    .create();
  Logger.log("Installed onFormSubmit trigger for form %s", formId);

  const responsesSheetId = PropertiesService.getScriptProperties().getProperty(RESPONSES_SHEET_ID_KEY);
  if (responsesSheetId) {
    ScriptApp.newTrigger("onFeedbackResponseReceived")
      .forSpreadsheet(responsesSheetId)
      .onFormSubmit()
      .create();
    Logger.log("Installed onFeedbackResponseReceived trigger for spreadsheet %s", responsesSheetId);
  }
}

/**
 * Remove stale ScriptProperties entries for forms and sheets that no longer exist.
 * Safe to run at any time — active entries are left untouched.
 *
 * USAGE:
 *   clasp run 'cleanup'
 */
/**
 * Returns true if the form file is trashed or inaccessible.
 * Uses the Drive API v3 directly so the result is authoritative regardless of
 * which Google account owns the Trash (DriveApp.isTrashed() is user-scoped).
 */
function isFormStale_(formId) {
  try {
    const resp = UrlFetchApp.fetch(
      "https://www.googleapis.com/drive/v3/files/" + formId + "?fields=trashed",
      { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return true; // 403/404 → inaccessible
    return JSON.parse(resp.getContentText()).trashed === true;
  } catch (e) {
    return true;
  }
}

function cleanup() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const toDelete = [];

  const responsesSheetId = allProps[RESPONSES_SHEET_ID_KEY];
  const ss = responsesSheetId ? SpreadsheetApp.openById(responsesSheetId) : null;
  const existingSheets = ss ? ss.getSheets() : [];
  const existingSheetIds = new Set(existingSheets.map(s => String(s.getSheetId())));

  Object.keys(allProps).forEach(key => {
    if (key.startsWith(PROPERTY_PREFIX + "notify.")) {
      const sheetId = key.slice((PROPERTY_PREFIX + "notify.").length);
      if (!existingSheetIds.has(sheetId)) {
        toDelete.push(key);
        Logger.log("Removing stale notify entry for missing sheet %s", sheetId);
      } else {
        // Sheet exists — check if its form is still properly linked to this spreadsheet.
        // A form in Trash is still "accessible" but stale; a permanently deleted form throws.
        const [, formId] = allProps[key].split("|");  // format: email|formId|formTitle
        let stale = false;
        if (isFormStale_(formId)) {
          stale = true;
        } else {
          try {
            const form = FormApp.openById(formId);
            stale = form.getDestinationType() !== FormApp.DestinationType.SPREADSHEET
              || form.getDestinationId() !== responsesSheetId;
          } catch (e) {
            stale = true;
          }
        }
        if (stale) {
          toDelete.push(key);
          Logger.log("Removing stale notify entry for form %s", formId);
          const sheet = existingSheets.find(s => String(s.getSheetId()) === sheetId);
          if (sheet) {
            try { FormApp.openById(formId).removeDestination(); } catch (e) {}
            try {
              ss.deleteSheet(sheet);
              Logger.log("Deleted orphaned response sheet %s", sheetId);
            } catch (e) {
              Logger.log("Could not delete orphaned sheet %s: %s", sheetId, e);
            }
          }
        }
      }
    } else if (key.startsWith(PROPERTY_PREFIX + "pendingNotify.")) {
      const formId = key.slice((PROPERTY_PREFIX + "pendingNotify.").length);
      let stale = false;
      if (isFormStale_(formId)) {
        stale = true;
      } else {
        try {
          const form = FormApp.openById(formId);
          stale = form.getDestinationType() !== FormApp.DestinationType.SPREADSHEET
            || form.getDestinationId() !== responsesSheetId;
        } catch (e) {
          stale = true;
        }
      }
      if (stale) {
        toDelete.push(key);
        Logger.log("Removing stale pendingNotify entry for form %s", formId);
        if (ss) {
          const linkedSheet = existingSheets.find(s => {
            const url = s.getFormUrl();
            return url && url.includes(formId);
          });
          if (linkedSheet) {
            try { FormApp.openById(formId).removeDestination(); } catch (e) {}
            try {
              ss.deleteSheet(linkedSheet);
              Logger.log("Deleted orphaned sheet '%s' for pending form %s", linkedSheet.getName(), formId);
            } catch (e) {
              Logger.log("Could not delete sheet for pending form %s: %s", formId, e);
            }
          }
        }
      }
    }
  });

  toDelete.forEach(key => props.deleteProperty(key));
  Logger.log("Cleanup complete. Removed %d stale entries.", toDelete.length);

  // Fallback: scan all sheets directly for orphaned form links.
  // Catches sheets with no corresponding ScriptProperty entry (e.g. manually cleaned props,
  // or forms generated before property tracking was in place).
  if (ss) {
    ss.getSheets().forEach(sheet => {
      const formUrl = sheet.getFormUrl();
      if (!formUrl) return;
      const match = formUrl.match(/\/forms\/d\/([^/]+)\//);
      const linkedFormId = match && match[1];
      if (!linkedFormId || !isFormStale_(linkedFormId)) return;
      // Unlink first (automation is owner, so removeDestination always succeeds).
      try { FormApp.openById(linkedFormId).removeDestination(); } catch (e) {}
      try {
        ss.deleteSheet(sheet);
        Logger.log("Deleted orphaned sheet '%s'", sheet.getName());
      } catch (e) {
        Logger.log("Could not delete sheet '%s': %s", sheet.getName(), e);
      }
    });
  }
}

/**
 * Remove all triggers installed by this project.
 *
 * USAGE:
 *   clasp run 'uninstall'
 */
function uninstall() {
  const managed = ["onFormSubmit", "onFeedbackResponseReceived"];
  ScriptApp.getProjectTriggers()
    .filter(t => managed.includes(t.getHandlerFunction()))
    .forEach(t => {
      ScriptApp.deleteTrigger(t);
      Logger.log("Uninstalled trigger for %s", t.getHandlerFunction());
    });
}

/**
 * Fires immediately when someone submits a response to any generated feedback form.
 * Triggered via the responses spreadsheet (push-based, no polling).
 * All generated forms share one spreadsheet; each form gets its own tab.
 *
 * Fast path: notify.<sheetId> already resolved (all responses after the first).
 * Resolve path: first response — use sheet.getFormUrl() to find the linked form,
 *   match against pendingNotify.<formId>, cache the result, then notify.
 */
function onFeedbackResponseReceived(e) {
  const sheet = e.range.getSheet();
  const sheetId = sheet.getSheetId();
  const notifyKey = PROPERTY_PREFIX + "notify." + sheetId;
  const props = PropertiesService.getScriptProperties();

  let value = props.getProperty(notifyKey);

  if (!value) {
    // First response for this sheet — resolve via the sheet's linked form URL.
    const formUrl = sheet.getFormUrl();
    const match = formUrl && formUrl.match(/\/forms\/d\/([^/]+)\//);
    const linkedFormId = match && match[1];

    if (linkedFormId) {
      const pendingKey = PROPERTY_PREFIX + "pendingNotify." + linkedFormId;
      const pendingValue = props.getProperty(pendingKey);
      if (pendingValue) {
        const [coordinatorEmail, formTitle] = pendingValue.split("|");
        value = coordinatorEmail + "|" + linkedFormId + "|" + formTitle;
        props.setProperty(notifyKey, value);
        props.deleteProperty(pendingKey);
        Logger.log("Resolved pending notification for form %s → sheet %s", linkedFormId, sheetId);
      }
    }
  }

  if (!value) {
    Logger.log("No coordinator info found for sheet %s (%s), skipping", sheetId, sheet.getName());
    return;
  }

  const parts = value.split("|");
  const coordinatorEmail = parts[0];
  const formId = parts[1];
  const formTitle = parts[2] || "a feedback form";
  const editUrl = "https://docs.google.com/forms/d/" + formId + "/edit#responses";

  MailApp.sendEmail({
    to: coordinatorEmail,
    subject: "New feedback response received: " + formTitle,
    htmlBody: `Someone submitted a response to <a href="${editUrl}">${formTitle}</a>.`
  });
}

/**
 * Triggered on form submit
 */
function onFormSubmit(e) {
  Logger.log(JSON.stringify(e.response, null, 2));

  const formResponse = e.response;  // Google Form Response object

  const email = formResponse.getRespondentEmail();

  // Get all answers
  const itemResponses = formResponse.getItemResponses();

  const answers = {};
  itemResponses.forEach(item => {
    answers[item.getItem().getTitle()] = item.getResponse();
  });

  Logger.log(email);
  Logger.log(answers);

  const employeeName = answers["Name"];
  let roles = ["generic"];

  let formRoles = answers["Roles"];
  if (!Array.isArray(formRoles)) {
    // maybe a single string, split by comma
    formRoles = formRoles.split(",").map(r => r.trim());
  } else {
    formRoles = formRoles.map(r => r.trim());
  }

  roles = roles.concat(formRoles);

  const customQuestions = answers["Custom Questions"] || "";

  const templates = fetchTemplatesForRoles(roles);
  const mergedSections = mergeTemplates(templates);

  const form = createFeedbackForm(employeeName, customQuestions, mergedSections);

  const formFile = DriveApp.getFileById(form.getId());

  const folderId = PropertiesService.getScriptProperties().getProperty(FOLDER_ID_KEY);
  if (folderId) {
    formFile.moveTo(DriveApp.getFolderById(folderId));
  }

  // Link form to the shared responses spreadsheet and register for push notifications.
  // Resolution of the sheet tab is deferred to the first response (onFeedbackResponseReceived)
  // via sheet.getFormUrl(), avoiding unreliable SpreadsheetApp caching issues.
  const responsesSheetId = PropertiesService.getScriptProperties().getProperty(RESPONSES_SHEET_ID_KEY);
  if (responsesSheetId) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, responsesSheetId);
    PropertiesService.getScriptProperties().setProperty(
      PROPERTY_PREFIX + "pendingNotify." + form.getId(), email + "|Feedback for " + employeeName
    );
  }

  form.addEditor(email);

  MailApp.sendEmail({
    to: email,
    subject: "Your feedback form is ready",
    htmlBody: `Your feedback form is ready!<br><br>
              You can edit it from <a href="${form.getEditUrl()}">here</a>.<br><br>
              And use this link to send it to others: <a href="${form.getPublishedUrl()}">${form.getPublishedUrl()}</a>.`
  });
}

/**
 * Fetch the role-to-filename mapping from the feedback-templates repo.
 *
 * Expected format in roles.json:
 *   { "Human Readable Role": "filename-stem", ... }
 *
 * The keys must match the option values in the intake form's Roles field.
 * The values are filenames in the repo without the .md extension.
 *
 * Example:
 *   {
 *     "generic":             "generic",
 *     "Software Engineer":   "engineer",
 *     "Engineering Manager": "manager"
 *   }
 */
function fetchRolesMapping() {
  const url = GITHUB_RAW_BASE + ROLES_MANIFEST_FILE;
  Logger.log("Fetching roles mapping from %s", url);
  const resp = UrlFetchApp.fetch(url);
  return JSON.parse(resp.getContentText());
}

/**
 * Fetch Markdown templates for selected roles using the roles mapping.
 * Roles not present in roles.json are skipped with a warning.
 */
function fetchTemplatesForRoles(roles) {
  const mapping = fetchRolesMapping();
  const templates = {};

  roles.forEach(role => {
    const filename = mapping[role];
    if (!filename) {
      Logger.log("No mapping found for role '%s' in %s, skipping", role, ROLES_MANIFEST_FILE);
      return;
    }

    const url = GITHUB_RAW_BASE + TEMPLATES_FOLDER + filename + ".md";
    Logger.log("Fetching template for role '%s' from %s", role, url);

    try {
      const resp = UrlFetchApp.fetch(url);
      templates[role] = resp.getContentText();
    } catch (e) {
      Logger.log("Error fetching template for role '%s': %s", role, e);
    }
  });
  return templates;
}

/**
 * Merge templates into sections & questions
 */
function mergeTemplates(templates) {
  const sections = {};
  for (const role in templates) {
    const text = templates[role];
    const blocks = text.split("\n# "); // split sections by Markdown header
    blocks.forEach(block => {
      const lines = block.split("\n").map(l => l.trim()).filter(l => l);
      if (lines.length === 0) return;
      let sectionName = lines[0].replace(/^#\s*/, "");
      let questions = lines.slice(1).map(l => l.replace(/^- /, ""));
      if (!sections[sectionName]) sections[sectionName] = [];
      questions.forEach(q => {
        if (!sections[sectionName].includes(q)) sections[sectionName].push(q);
      });
    });
  }
  return sections;
}

function createFeedbackForm(employeeName, customQuestions, sections) {
  const form = FormApp.create(`Feedback for ${employeeName}`);
  form.setDescription("Please provide your feedback using the prompts below.");

  // Add custom questions at the top
  if (customQuestions) {
    createSection(form, "Custom Questions", customQuestions.split("\n"));
  }

  // Add template sections
  Object.keys(sections).forEach(sectionTitle => {
    createSection(form, sectionTitle, sections[sectionTitle]);
  });

  return form;
}

function createSection(form, name, questions) {
  var section_item = form.addPageBreakItem();
  section_item.setTitle(name);

  var title_and_description_item = form.addSectionHeaderItem().setTitle("For each question, please include an Observation, Impact and Question/Advice");

  questions.forEach(q => {
    addGroupedQuestion(form, q);
  });
}

/**
 * Add Observation / Impact / Question fields for a single question
 */
function addGroupedQuestion(form, questionText) {
  // Single paragraph with all prompts inside
  form.addParagraphTextItem()
    .setTitle(questionText)
    .setRequired(false);
}
