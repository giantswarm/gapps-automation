const GITHUB_RAW_BASE = "https://raw.githubusercontent.com/giantswarm/feedback-templates/refs/heads/main/";

const ROLES_MANIFEST_FILE = "roles.json";

const TEMPLATES_FOLDER = "questions/";

const PROPERTY_PREFIX = "EmployeeFeedback.";

const FORM_ID_KEY = PROPERTY_PREFIX + "formId";

const FOLDER_ID_KEY = PROPERTY_PREFIX + "folderId";

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
}

/**
 * Remove all onFormSubmit triggers installed by this project.
 *
 * USAGE:
 *   clasp run 'uninstall'
 */
function uninstall() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onFormSubmit")
    .forEach(t => {
      ScriptApp.deleteTrigger(t);
      Logger.log("Uninstalled trigger for %s", t.getHandlerFunction());
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

  formFile.setOwner(email);

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
