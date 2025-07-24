// Global variables to store parsed data and headers
let workbook = null;
let sheetData = []; // Stores all data from the first sheet (including headers as first row)
let headers = [];   // Stores the first row (headers)
let availableFormFields = []; // Stores fields found on the active tab
let groupedFormFields = {}; // Stores fields grouped by their 'surroundingText' or derived name
let learnedMappings = {}; // Stores user's preferred mappings for schema learning

// DOM Elements - Tabs
const autoFillTab = document.getElementById("autoFillTab");
const autoClickTab = document.getElementById("autoClickTab");
const autoFillSection = document.getElementById("autoFillSection");
const autoClickSection = document.getElementById("autoClickSection");

// DOM Elements - Auto Fill Section
const fileInput = document.getElementById('fileInput');
const dropArea = document.getElementById('drop-area'); // New: Reference to the drag and drop area
const fileNameDisplay = document.getElementById('fileNameDisplay'); // Still used for "Drag & Drop..." text
const fileStatusMessage = document.getElementById('fileStatusMessage'); // New: The combined status div
const fileMessage = document.getElementById('fileMessage'); // New: Span inside status div for text
const fileStatusIcon = document.getElementById('fileStatusIcon'); // New: Icon inside status div
const dataDisplaySection = document.getElementById('dataDisplaySection');
const headersDisplay = document.getElementById('headersDisplay');
const scanFieldsButton = document.getElementById('scanFieldsButton');
const scanMessage = document.getElementById('scanMessage');
const fieldMappingSection = document.getElementById('fieldMappingSection');
const mappingContainer = document.getElementById('mappingContainer');
const fillDataButton = document.getElementById('fillDataButton');
const fillDataMessage = document.getElementById('fillDataMessage');
const fillEmptyOnlyCheckbox = document.getElementById('fillEmptyOnlyCheckbox');
const oneClickAutofillButton = document.getElementById('oneClickAutofillButton');
const testFillButton = document.getElementById('testFillButton');
const previewValuesButton = document.getElementById('previewValuesButton');
const testFillMessage = document.getElementById('testFillMessage');
const previewValuesMessage = document.getElementById('previewValuesMessage');

// DOM Elements - Auto Click Section
const scanButtons = document.getElementById("scanButtons");
const clickableButtonsContainer = document.getElementById("clickableButtonsContainer");
const clickCountInput = document.getElementById("clickCount");
const startClickingButton = document.getElementById("startClicking");
const autoClickMessage = document.getElementById("autoClickMessage");
const selectButtonCard = document.getElementById("selectButtonCard"); // New card for step 2
const clickControlSection = document.getElementById("clickControlSection"); // Card for step 3

// DOM Elements - Theme Toggle
const themeToggle = document.getElementById('themeToggle');


// --- Event Listeners ---

// Tab switching logic
autoFillTab.addEventListener("click", () => switchTab('autoFill'));
autoClickTab.addEventListener("click", () => switchTab('autoClick'));

// Theme toggle logic
themeToggle.addEventListener('change', () => {
  document.body.classList.toggle('dark', themeToggle.checked);
  // Also toggle dark class on the main container and relevant cards
  document.querySelector('.main-container').classList.toggle('dark', themeToggle.checked);
  document.querySelectorAll('.card').forEach(card => card.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.section-heading').forEach(heading => heading.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.card-heading').forEach(heading => heading.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.action-button').forEach(button => button.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.form-input').forEach(input => input.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.drop-area').forEach(drop => drop.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.browse-button').forEach(button => button.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.file-status').forEach(status => status.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.headers-list').forEach(list => list.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.headers-list span').forEach(span => span.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.checkbox-label').forEach(label => label.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.radio-label').forEach(label => label.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.message-box').forEach(box => box.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.auto-mapped-badge').forEach(badge => badge.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.tab-buttons').forEach(buttons => buttons.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.tab-button').forEach(button => button.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('#clickableButtonsContainer').forEach(container => container.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('#clickableButtonsContainer > div').forEach(div => div.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('#clickableButtonsContainer label').forEach(label => label.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('#clickControlSection').forEach(section => section.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('#clickControlSection label').forEach(label => label.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.selected-button-highlight').forEach(highlight => highlight.classList.toggle('dark', themeToggle.checked));
  document.querySelectorAll('.slider').forEach(slider => slider.classList.toggle('dark', themeToggle.checked));


  localStorage.setItem('theme', themeToggle.checked ? 'dark' : 'light');
});

// File input change
fileInput.addEventListener('change', handleFile);

// Drag and drop functionality for file input
dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.classList.add('border-blue-600', 'bg-blue-100', 'dark:border-indigo-500', 'dark:bg-indigo-900');
    fileNameDisplay.classList.add('group-hover:text-blue-700', 'dark:group-hover:text-blue-300'); // Add hover text color
});
dropArea.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropArea.classList.remove('border-blue-600', 'bg-blue-100', 'dark:border-indigo-500', 'dark:bg-indigo-900');
    fileNameDisplay.classList.remove('group-hover:text-blue-700', 'dark:group-hover:text-blue-300');
});
dropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dropArea.classList.remove('border-blue-600', 'bg-blue-100', 'dark:border-indigo-500', 'dark:bg-indigo-900');
    fileNameDisplay.classList.remove('group-hover:text-blue-700', 'dark:group-hover:text-blue-300');
    if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
        handleFile();
    }
});

// Scan fields button click
scanFieldsButton.addEventListener('click', async () => {
    await scanCurrentTabFields();
});

// Fill data button click
fillDataButton.addEventListener('click', () => fillDataInTab(false));

// One-Click Autofill button click
oneClickAutofillButton.addEventListener('click', () => fillDataInTab(true));

// Test Fill button click
testFillButton.addEventListener('click', () => testFillFirstRow());

// Preview Values button click
previewValuesButton.addEventListener('click', () => previewMappedValues());

// Auto Click Event Listeners
scanButtons.addEventListener("click", async () => {
    displayFileStatusMessage('<i class="fas fa-spinner fa-spin"></i>Scanning for clickable buttons...', 'info', autoClickMessage, true);
    scanButtons.disabled = true;
    scanButtons.innerHTML = 'Scanning... <span class="loading-spinner"></span>';
    clickableButtonsContainer.innerHTML = '<p class="text-gray-500 dark:text-gray-400 text-sm p-2">Scanning...</p>'; // Clear previous buttons and show scanning message

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayFileStatusMessage('Could not get active tab.', 'error', autoClickMessage, true);
            return;
        }

        const response = await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            files: ['content.js']
        }).then(() => {
            return chrome.tabs.sendMessage(tab.id, { action: "scanClickables" });
        });

        if (response && response.clickables) {
            if (response.clickables.length > 0) {
                clickableButtonsContainer.innerHTML = ''; // Clear scanning message
                response.clickables.forEach((btn, index) => {
                    const btnDiv = document.createElement("div");
                    btnDiv.className = "flex items-center mb-1 last:mb-0 border-b border-gray-200 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-md transition-colors duration-200 p-2"; // Add styling classes
                    btnDiv.innerHTML = `
                        <input type="radio" name="clickable" value="${btn.stableId}" id="btn-${index}" class="mr-2 flex-shrink-0">
                        <label for="btn-${index}" class="text-sm text-gray-700 dark:text-gray-300 flex-grow">${btn.text}</label>
                    `;
                    clickableButtonsContainer.appendChild(btnDiv);

                    // Add event listener to highlight selected radio button
                    btnDiv.querySelector('input[type="radio"]').addEventListener('change', (e) => {
                        document.querySelectorAll('#clickableButtonsContainer > div').forEach(div => {
                            div.classList.remove('selected-button-highlight', 'dark:selected-button-highlight');
                        });
                        if (e.target.checked) {
                            btnDiv.classList.add('selected-button-highlight', document.body.classList.contains('dark') ? 'dark:selected-button-highlight' : '');
                        }
                    });
                });
                selectButtonCard.classList.remove("hidden"); // Show Step 2 card
                clickControlSection.classList.remove("hidden"); // Show Step 3 card
                displayFileStatusMessage(`<i class="fas fa-check-circle"></i>Found ${response.clickables.length} clickable elements.`, 'success', autoClickMessage, true);
            } else {
                clickableButtonsContainer.innerHTML = "<p class='text-gray-500 dark:text-gray-400 text-sm p-2'>No clickable buttons found on this page.</p>";
                selectButtonCard.classList.add("hidden");
                clickControlSection.classList.add("hidden");
                displayFileStatusMessage('<i class="fas fa-info-circle"></i>No clickable buttons found on the current tab.', 'info', autoClickMessage, true);
            }
        } else {
            clickableButtonsContainer.innerHTML = "<p class='text-gray-500 dark:text-gray-400 text-sm p-2'>Failed to scan for buttons.</p>";
            selectButtonCard.classList.add("hidden");
            clickControlSection.classList.add("hidden");
            displayFileStatusMessage('<i class="fas fa-exclamation-triangle"></i>Failed to get clickable elements from the current tab.', 'error', autoClickMessage, true);
        }
    } catch (error) {
        console.error("Error scanning buttons:", error);
        displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error scanning buttons: ${error.message}.`, 'error', autoClickMessage, true);
        clickableButtonsContainer.innerHTML = "<p class='text-gray-500 dark:text-gray-400 text-sm p-2'>Error scanning for buttons. Check console for details.</p>";
        selectButtonCard.classList.add("hidden");
        clickControlSection.classList.add("hidden");
    } finally {
        scanButtons.disabled = false;
        scanButtons.innerHTML = '<i class="fas fa-sync-alt mr-2"></i>Scan for Clickable Buttons';
    }
});

startClickingButton.addEventListener("click", async () => {
    const count = parseInt(clickCountInput.value);
    const selected = document.querySelector('input[name="clickable"]:checked');
    const stableId = selected?.value;

    if (isNaN(count) || count < 1) {
        displayFileStatusMessage("Please enter a valid number of clicks (1 or more).", "error", autoClickMessage, false);
        return;
    }

    if (!stableId) {
        displayFileStatusMessage("Please select a button to click.", "error", autoClickMessage, false);
        return;
    }

    displayFileStatusMessage(`<i class="fas fa-circle-notch fa-spin"></i>Attempting to click ${count} time(s)...`, 'info', autoClickMessage, true);
    startClickingButton.disabled = true;
    startClickingButton.innerHTML = 'Clicking... <span class="loading-spinner"></span>';


    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayFileStatusMessage('Could not get active tab.', 'error', autoClickMessage, false);
            return;
        }

        const response = await chrome.tabs.sendMessage(tab.id, {
            action: "performClick",
            stableId,
            count
        });

        if (response?.status === "success") {
            displayFileStatusMessage(`<i class="fas fa-check-circle"></i>Clicked ${count} time(s) successfully.`, "success", autoClickMessage, true);
        } else {
            console.error(`Error during click: ${response?.message || 'Unknown error.'}`);
            displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Failed to click: ${response?.message || 'Unknown error.'}`, "error", autoClickMessage, true);
        }
    } catch (error) {
        console.error("Error performing click:", error);
        displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error performing click: ${error.message}.`, "error", autoClickMessage, true);
    } finally {
        startClickingButton.disabled = false;
        startClickingButton.innerHTML = '<i class="fas fa-bullseye mr-2"></i>Start Clicking';
    }
});

// Load learned mappings and initial tab on startup
document.addEventListener('DOMContentLoaded', () => {
    loadLearnedMappings();
    loadTabPreference();
    // Load theme preference on DOMContentLoaded
    const theme = localStorage.getItem('theme');
    if (theme === 'dark') {
      document.body.classList.add('dark');
      // Apply dark class to other elements as well for initial load
      document.querySelector('.main-container').classList.add('dark');
      document.querySelectorAll('.card').forEach(card => card.classList.add('dark'));
      document.querySelectorAll('.section-heading').forEach(heading => heading.classList.add('dark'));
      document.querySelectorAll('.card-heading').forEach(heading => heading.classList.add('dark'));
      document.querySelectorAll('.action-button').forEach(button => button.classList.add('dark'));
      document.querySelectorAll('.form-input').forEach(input => input.classList.add('dark'));
      document.querySelectorAll('.drop-area').forEach(drop => drop.classList.add('dark'));
      document.querySelectorAll('.browse-button').forEach(button => button.classList.add('dark'));
      document.querySelectorAll('.file-status').forEach(status => status.classList.add('dark'));
      document.querySelectorAll('.headers-list').forEach(list => list.classList.add('dark'));
      document.querySelectorAll('.headers-list span').forEach(span => span.classList.add('dark'));
      document.querySelectorAll('.checkbox-label').forEach(label => label.classList.add('dark'));
      document.querySelectorAll('.radio-label').forEach(label => label.classList.add('dark'));
      document.querySelectorAll('.message-box').forEach(box => box.classList.add('dark'));
      document.querySelectorAll('.auto-mapped-badge').forEach(badge => badge.classList.add('dark'));
      document.querySelectorAll('.tab-buttons').forEach(buttons => buttons.classList.add('dark'));
      document.querySelectorAll('.tab-button').forEach(button => button.classList.add('dark'));
      document.querySelectorAll('#clickableButtonsContainer').forEach(container => container.classList.add('dark'));
      document.querySelectorAll('#clickableButtonsContainer > div').forEach(div => div.classList.add('dark'));
      document.querySelectorAll('#clickableButtonsContainer label').forEach(label => label.classList.add('dark'));
      document.querySelectorAll('#clickControlSection').forEach(section => section.classList.add('dark'));
      document.querySelectorAll('#clickControlSection label').forEach(label => label.classList.add('dark'));
      document.querySelectorAll('.selected-button-highlight').forEach(highlight => highlight.classList.add('dark'));
      document.querySelectorAll('.slider').forEach(slider => slider.classList.add('dark'));
      themeToggle.checked = true;
    }
});


// --- Functions ---

/**
 * Switches between the Auto Fill and Auto Click tabs.
 * @param {string} activeTabId - The ID of the tab to activate ('autoFill' or 'autoClick').
 */
async function switchTab(activeTabId) {
    // Remove active class from all tabs and add to the selected one
    autoFillTab.classList.remove('active', 'dark:active');
    autoClickTab.classList.remove('active', 'dark:active');
    
    const targetTabButton = document.getElementById(`${activeTabId}Tab`);
    targetTabButton.classList.add('active');
    if (document.body.classList.contains('dark')) {
        targetTabButton.classList.add('dark:active');
    }

    // Hide all tab sections and show the selected one
    autoFillSection.classList.add('hidden');
    autoClickSection.classList.add('hidden');
    document.getElementById(`${activeTabId}Section`).classList.remove('hidden');

    // Save tab preference
    try {
        await chrome.storage.sync.set({ activeTab: activeTabId });
    } catch (error) {
        console.error('Error saving tab preference:', error);
    }
}

/**
 * Loads the last active tab preference from chrome.storage.sync.
 */
async function loadTabPreference() {
    try {
        const result = await chrome.storage.sync.get('activeTab');
        const lastActiveTab = result.activeTab || 'autoFill'; // Default to autoFill
        switchTab(lastActiveTab);
    } catch (error) {
        console.error('Error loading tab preference:', error);
        switchTab('autoFill'); // Fallback to default
    }
}


/**
 * Handles the file selection and reads its content.
 */
function handleFile() {
    const file = fileInput.files[0];
    if (!file) {
        displayFileStatusMessage('<i class="fas fa-exclamation-triangle"></i>No file selected.', 'error', fileStatusMessage, true);
        return;
    }

    fileNameDisplay.textContent = `File: "${file.name}"`; // Update text to show file name
    displayFileStatusMessage(`<i class="fas fa-spinner fa-spin"></i>Reading "${file.name}"...`, 'info', fileStatusMessage, true);

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            // Read the workbook
            workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet name
            const sheetName = workbook.SheetNames[0];
            // Convert the first sheet to JSON, ensuring header:1 to get raw array of arrays
            sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

            if (sheetData.length === 0) {
                displayFileStatusMessage('<i class="fas fa-exclamation-triangle"></i>The selected file is empty or could not be parsed.', 'error', fileStatusMessage, true);
                dataDisplaySection.classList.add('hidden');
                oneClickAutofillButton.classList.add('hidden'); // Hide autofill if no data
                return;
            }

            // The first row is the headers
            headers = sheetData[0];
            
            displayHeaders(headers);
            dataDisplaySection.classList.remove('hidden');
            displayFileStatusMessage(`<i class="fas fa-check-circle"></i>File "${file.name}" loaded successfully.`, 'success', fileStatusMessage, true);

            // If fields were already scanned, re-setup mapping with new headers and re-auto-map
            if (Object.keys(groupedFormFields).length > 0) { // Check groupedFormFields for existing scan
                setupFieldMapping(groupedFormFields, headers); // Re-setup mapping with new headers
                autoMapFields(groupedFormFields, headers); // Re-run auto-map
                fieldMappingSection.classList.remove('hidden');
                oneClickAutofillButton.classList.remove('hidden'); // Show autofill if fields and headers exist
            } else {
                // If no fields scanned yet, hide autofill button
                oneClickAutofillButton.classList.add('hidden');
            }

        } catch (error) {
            console.error("Error reading file:", error);
            displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error reading file: ${error.message}. Please ensure it's a valid spreadsheet format.`, 'error', fileStatusMessage, true);
            dataDisplaySection.classList.add('hidden');
            oneClickAutofillButton.classList.add('hidden');
        }
    };

    reader.onerror = function(e) {
        console.error("FileReader error:", e);
        displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error reading file: ${e.target.error.name}.`, 'error', fileStatusMessage, true);
        dataDisplaySection.classList.add('hidden');
        oneClickAutofillButton.classList.add('hidden');
    };

    reader.readAsArrayBuffer(file);
}

/**
 * Displays the extracted headers in the UI.
 * @param {string[]} headersArray - Array of header strings.
 */
function displayHeaders(headersArray) {
    headersDisplay.innerHTML = ''; // Clear previous headers
    if (headersArray.length > 0) {
        headersArray.forEach(header => {
            const span = document.createElement('span');
            span.className = 'bg-indigo-100 text-indigo-800 px-5 py-2 rounded-full text-base font-medium shadow-sm flex items-center transition-colors duration-200 hover:bg-indigo-200 cursor-default dark:bg-indigo-700 dark:text-indigo-100 dark:hover:bg-indigo-600';
            span.textContent = header;
            headersDisplay.appendChild(span);
        });
    } else {
        headersDisplay.textContent = 'No headers found in the first row.';
        headersDisplay.classList.add('text-gray-500', 'dark:text-gray-400', 'text-sm');
    }
}

/**
 * Sends a message to the content script to scan for form fields.
 */
async function scanCurrentTabFields() {
    if (!headers || headers.length === 0) {
        displayMessage(scanMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please upload a file with headers first.', 'error', true);
        return;
    }

    displayMessage(scanMessage, '<i class="fas fa-spinner fa-spin mr-2"></i>Scanning current tab for fields...', 'info', true);
    scanFieldsButton.disabled = true; // Disable button during scan
    scanFieldsButton.innerHTML = 'Scanning... <span class="loading-spinner"></span>';
    oneClickAutofillButton.classList.add('hidden'); // Hide autofill button during scan

    try {
        // Get the active tab
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(scanMessage, 'Could not get active tab.', 'error', false);
            return;
        }

        // Inject content script if not already injected and send message to scan
        const response = await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            files: ['content.js'] // Inject content.js
        }).then(() => {
            // Once injected, send a message to the content script to perform the scan
            return chrome.tabs.sendMessage(tab.id, { action: 'scanFields' });
        });

        if (response && response.fields) {
            availableFormFields = response.fields;
            if (availableFormFields.length > 0) {
                groupedFormFields = groupFieldsBySignature(availableFormFields); // Use smarter grouping
                setupFieldMapping(groupedFormFields, headers); // Call existing setup to build UI with grouped fields
                autoMapFields(groupedFormFields, headers); // Call auto-map to pre-select dropdowns for groups
                fieldMappingSection.classList.remove('hidden');
                displayMessage(scanMessage, `<i class="fas fa-check-circle mr-2"></i>Found ${availableFormFields.length} fields on the page, grouped into ${Object.keys(groupedFormFields).length} sections. Attempting auto-mapping.`, 'success', true);
                oneClickAutofillButton.classList.remove('hidden'); // Show autofill button if fields found
            } else {
                fieldMappingSection.classList.add('hidden');
                displayMessage(scanMessage, '<i class="fas fa-info-circle mr-2"></i>No input fields found on the current tab.', 'info', true);
                oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if no fields
            }
        } else {
            fieldMappingSection.classList.add('hidden');
            displayMessage(scanMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Failed to get fields from the current tab.', 'error', true);
            oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if error
        }
    } catch (error) {
        console.error("Error scanning fields:", error);
        displayMessage(scanMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error scanning fields: ${error.message}.`, 'error', true);
        fieldMappingSection.classList.add('hidden');
        oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if error
    } finally {
        scanFieldsButton.disabled = false;
        scanFieldsButton.innerHTML = '<i class="fas fa-sync-alt mr-2"></i>Scan Current Tab for Fields';
    }
}

/**
 * Generates a comprehensive signature for a form field using multiple attributes.
 * This signature is used for smarter grouping and fuzzy matching.
 * @param {Object} field - The field object from content.js.
 * @returns {string} A combined string representing the field's unique signature.
 */
function generateFieldSignature(field) {
    const parts = [
        field.labelText,
        field.name,
        field.placeholder,
        field.ariaLabel,
        field.title,
        field.autocomplete,
        field.surroundingText
    ].filter(Boolean)
     .map(str => str.toLowerCase().replace(/[^a-z0-9\s]/g, '').trim())
     .filter(str => str.length > 1);

    return [...new Set(parts)].join(" | ");
}

/**
 * Groups form fields by their generated signature to create logical sections.
 * @param {Array<Object>} formFields - Array of field objects from content.js.
 * @returns {Object} An object where keys are grouping contexts (signatures) and values are arrays of fields.
 */
function groupFieldsBySignature(formFields) {
    const groups = {};
    formFields.forEach(field => {
        const signature = generateFieldSignature(field);
        const groupIdentifier = signature || field.htmlId || 'Unnamed Field Group';
        
        if (!groups[groupIdentifier]) {
            groups[groupIdentifier] = [];
        }
        groups[groupIdentifier].push(field);
    });
    return groups;
}

/**
 * Sets up the mapping section with dropdowns for each grouped form field context.
 * @param {Object} groupedFields - Object of grouped fields (e.g., { 'Context A': [field1, field2], 'Context B': [field3] }).
 * @param {string[]} headersArray - Array of header strings for dropdown options.
 */
function setupFieldMapping(groupedFields, headersArray) {
    mappingContainer.innerHTML = ''; // Clear previous mappings

    const groupKeys = Object.keys(groupedFields);
    if (groupKeys.length === 0) {
        mappingContainer.textContent = 'No mappable field groups found on the page.';
        mappingContainer.classList.add('text-gray-500', 'dark:text-gray-400', 'text-sm');
        return;
    }
    mappingContainer.classList.remove('text-gray-500', 'dark:text-gray-400', 'text-sm');

    groupKeys.forEach(contextKey => {
        const fieldsInGroup = groupedFields[contextKey];
        const mappingGroupItem = document.createElement('div');
        mappingGroupItem.className = 'field-mapping-group-item card'; // Apply card styling
        if (document.body.classList.contains('dark')) {
            mappingGroupItem.classList.add('dark');
        }

        const groupHeader = document.createElement('h3');
        groupHeader.className = 'text-md font-semibold text-gray-800 dark:text-gray-200 mb-3 flex items-center card-heading';
        groupHeader.textContent = contextKey;
        if (document.body.classList.contains('dark')) {
            groupHeader.classList.add('dark');
        }
        mappingGroupItem.appendChild(groupHeader);

        const groupControl = document.createElement('div');
        groupControl.className = 'flex flex-col sm:flex-row items-start sm:items-center space-y-2 sm:space-y-0 sm:space-x-2 mb-3';

        const label = document.createElement('label');
        label.className = 'checkbox-label flex-shrink-0';
        if (document.body.classList.contains('dark')) {
            label.classList.add('dark');
        }

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'group-checkbox mr-2';
        checkbox.dataset.contextKey = contextKey;
        checkbox.checked = false;

        const span = document.createElement('span');
        span.className = 'text-sm text-gray-700 dark:text-gray-300';
        span.textContent = `Map fields for "${contextKey}"`;

        label.appendChild(checkbox);
        label.appendChild(span);
        groupControl.appendChild(label);

        const select = document.createElement('select');
        select.className = 'group-mapper flex-grow mt-2 sm:mt-0 form-input'; // Use form-input for styling
        select.dataset.contextKey = contextKey;
        select.disabled = true;
        if (document.body.classList.contains('dark')) {
            select.classList.add('dark');
        }

        const fragment = document.createDocumentFragment();
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = '-- Select Column --';
        fragment.appendChild(defaultOption);

        headersArray.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            fragment.appendChild(option);
        });
        select.appendChild(fragment);
        groupControl.appendChild(select);
        mappingGroupItem.appendChild(groupControl);

        // Display individual fields within this group (read-only)
        const fieldListContainer = document.createElement('div');
        fieldListContainer.className = 'mt-2 text-xs text-gray-600 dark:text-gray-400';
        fieldListContainer.innerHTML = '<p class="font-medium mb-1">Instances of this field on page:</p>';
        fieldsInGroup.forEach(field => {
            let displayName = field.labelText || field.name || field.placeholder || field.ariaLabel || field.title || 'Unnamed Field Instance';
            displayName = displayName.replace(/\[\d+\]/g, '').trim();

            if (displayName.startsWith('stable-id-') || displayName.startsWith('generated-id-')) {
                displayName = 'Unnamed Field Instance';
            }

            const fieldSpan = document.createElement('span');
            fieldSpan.className = 'inline-block bg-gray-100 text-gray-700 px-2 py-0.5 rounded-full mr-1 mb-1 dark:bg-gray-700 dark:text-gray-300';
            fieldSpan.textContent = displayName;
            fieldListContainer.appendChild(fieldSpan);
        });
        mappingGroupItem.appendChild(fieldListContainer);

        mappingContainer.appendChild(mappingGroupItem);

        // Event listener for group checkbox
        checkbox.addEventListener('change', (e) => {
            const currentSelect = mappingContainer.querySelector(`.group-mapper[data-context-key="${e.target.dataset.contextKey}"]`);
            currentSelect.disabled = !e.target.checked;
            if (!e.target.checked) {
                currentSelect.value = '';
                removeAutoMappedBadge(groupHeader); // Remove badge if unchecked
            } else {
                // If re-checked, try auto-mapping again for visual consistency
                autoMapFields(groupedFields, headersArray);
            }
        });

        // Event listener for select change to save mapping and update badge
        select.addEventListener('change', () => {
            saveLearnedMappings(); // Save whenever a mapping is changed
            updateBadgeForGroup(groupHeader, select.value, learnedMappings[contextKey]);
        });
    });
}

/**
 * Performs intelligent auto-mapping between grouped form fields and spreadsheet headers using Fuse.js.
 * Populates the mapping dropdowns and checks the corresponding checkboxes for groups.
 * @param {Object} groupedFields - Object of grouped fields.
 * @param {string[]} headersArray - Array of header strings.
 */
function autoMapFields(groupedFields, headersArray) {
    // Load learned mappings first
    loadLearnedMappings().then(() => {
        const fuseOptions = {
            includeScore: true,
            threshold: 0.4, // Lower is stricter, 0.4 allows for some flexibility
            keys: ['header']
        };

        const fuse = new Fuse(headersArray.map(h => ({ header: h })), fuseOptions);

        let autoMappedCount = 0;
        Object.keys(groupedFields).forEach(contextKey => {
            const checkbox = mappingContainer.querySelector(`.group-checkbox[data-context-key="${contextKey}"]`);
            const select = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
            const groupHeaderElement = checkbox.closest('.field-mapping-group-item').querySelector('h3');

            if (!checkbox || !select) return;

            let mappedType = 'unmapped'; // Default to unmapped

            // 1. Try to apply learned mapping first
            if (learnedMappings[contextKey] && headersArray.includes(learnedMappings[contextKey])) {
                checkbox.checked = true;
                select.disabled = false;
                select.value = learnedMappings[contextKey];
                mappedType = 'learned';
                autoMappedCount++;
                console.log(`Auto-mapping (Learned): "${contextKey}" -> "${learnedMappings[contextKey]}"`);
            } else {
                // 2. Fallback to fuzzy matching if no learned mapping or learned mapping no longer valid
                const result = fuse.search(contextKey)[0];
                if (result && result.score < 0.4) { // Confidence threshold
                    const bestMatchHeader = result.item.header;
                    checkbox.checked = true;
                    select.disabled = false;
                    select.value = bestMatchHeader;
                    mappedType = 'fuzzy';
                    autoMappedCount++;
                    console.log(`Auto-mapping (Fuzzy): "${contextKey}" -> "${bestMatchHeader}" (Score: ${result.score.toFixed(2)})`);
                } else {
                    // If not auto-mapped, ensure it's unchecked and disabled
                    checkbox.checked = false;
                    select.disabled = true;
                    select.value = '';
                    mappedType = 'unmapped';
                    console.log(`No strong auto-mapping for "${contextKey}" (Best score: ${result?.score.toFixed(2) || 'N/A'})`);
                }
            }
            updateBadgeForGroup(groupHeaderElement, mappedType, mappedType === 'fuzzy' ? result.score.toFixed(2) : '');
        });
        displayMessage(scanMessage, `<i class="fas fa-check-circle mr-2"></i>Auto-mapping complete. ${autoMappedCount} fields auto-mapped. Review and adjust if needed.`, 'success', true);
    }).catch(error => {
        console.error("Error loading learned mappings for auto-map:", error);
        displayMessage(scanMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error loading learned mappings: ${error.message}. Auto-mapping proceeded without them.`, 'error', true);
    });
}

/**
 * Adds or updates an "Auto-Matched" badge to the group header.
 * @param {HTMLElement} groupHeaderElement - The H3 element of the group.
 * @param {'learned'|'fuzzy'|'unmapped'} type - The type of mapping.
 * @param {string} [scoreText=''] - Optional score text for fuzzy matches.
 */
function updateBadgeForGroup(groupHeaderElement, type, scoreText = '') {
    let badge = groupHeaderElement.querySelector('.auto-mapped-badge');
    if (!badge) {
        badge = document.createElement('span');
        badge.className = 'auto-mapped-badge';
        groupHeaderElement.appendChild(badge);
    }
    
    // Remove previous type classes
    badge.classList.remove('learned', 'fuzzy', 'unmapped');
    badge.classList.add(type);
    if (document.body.classList.contains('dark')) {
        badge.classList.add('dark'); // Ensure dark mode class is applied to badge
    } else {
        badge.classList.remove('dark');
    }

    let badgeText = '';
    if (type === 'learned') {
        badgeText = 'Auto-Matched (Learned)';
    } else if (type === 'fuzzy') {
        badgeText = `Auto-Matched (Score: ${scoreText})`;
    } else {
        badgeText = 'Unmapped';
    }
    badge.textContent = badgeText;

    // Hide badge if unmapped or no selection
    if (type === 'unmapped' || !groupHeaderElement.parentElement.querySelector('.group-mapper').value) {
        badge.classList.add('hidden');
    } else {
        badge.classList.remove('hidden');
    }
}

/**
 * Removes the auto-mapped badge from a group header.
 * @param {HTMLElement} groupHeaderElement - The H3 element of the group.
 */
function removeAutoMappedBadge(groupHeaderElement) {
    const badge = groupHeaderElement.querySelector('.auto-mapped-badge');
    if (badge) {
        badge.classList.add('hidden');
    }
}

/**
 * Saves the currently selected mappings to chrome.storage.sync for future use.
 */
async function saveLearnedMappings() {
    const currentMappings = {};
    mappingContainer.querySelectorAll('.group-checkbox:checked').forEach(checkbox => {
        const contextKey = checkbox.dataset.contextKey;
        const selectElement = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
        const mappedColumnHeader = selectElement ? selectElement.value : '';
        if (mappedColumnHeader) {
            currentMappings[contextKey] = selectElement.value;
        }
    });

    try {
        await chrome.storage.sync.set({ learnedMappings: currentMappings });
        console.log('Learned mappings saved:', currentMappings);
        // After saving, re-run autoMapFields to update badges based on newly learned mappings
        if (Object.keys(groupedFormFields).length > 0 && headers.length > 0) {
            autoMapFields(groupedFormFields, headers);
        }
    } catch (error) {
        console.error('Error saving learned mappings:', error);
    }
}

/**
 * Loads learned mappings from chrome.storage.sync.
 */
async function loadLearnedMappings() {
    try {
        const result = await chrome.storage.sync.get('learnedMappings');
        learnedMappings = result.learnedMappings || {};
        console.log('Learned mappings loaded:', learnedMappings);
    } catch (error) {
        console.error('Error loading learned mappings:', error);
    }
}

/**
 * Sends a message to the content script to fill the fields on the active tab.
 * This function now iterates through spreadsheet rows and prepares a batch for filling.
 * @param {boolean} isAutoFill - True if triggered by the "One-Click Autofill" button.
 */
async function fillDataInTab(isAutoFill = false) {
    if (!workbook || sheetData.length <= 1) { // sheetData includes headers, so >1 means actual data rows
        displayMessage(fillDataMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please upload a spreadsheet file first.', 'error', true);
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(fillDataMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please scan for fields on the current tab first.', 'error', true);
        return;
    }

    const actualDataRows = sheetData.slice(1); // Get data rows, excluding headers

    const fillEmptyOnly = fillEmptyOnlyCheckbox.checked;
    
    // Collect all selected mappings: { contextKey: mappedColumnHeader }
    const selectedMappings = {};
    mappingContainer.querySelectorAll('.group-checkbox:checked').forEach(checkbox => {
        const contextKey = checkbox.dataset.contextKey;
        const selectElement = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
        const mappedColumnHeader = selectElement ? selectElement.value : '';
        if (mappedColumnHeader) {
            selectedMappings[contextKey] = mappedColumnHeader;
        }
    });

    if (Object.keys(selectedMappings).length === 0) {
        displayMessage(fillDataMessage, '<i class="fas fa-info-circle mr-2"></i>No fields selected for filling or no column mapped to selected groups.', 'error', true);
        return;
    }

    displayMessage(fillDataMessage, '<i class="fas fa-spinner fa-spin mr-2"></i>Preparing data for filling...', 'info', true);
    fillDataButton.disabled = true;
    fillDataButton.innerHTML = 'Filling... <span class="loading-spinner"></span>';
    oneClickAutofillButton.disabled = true;
    oneClickAutofillButton.innerHTML = 'One-Click Autofill... <span class="loading-spinner"></span>';
    testFillButton.disabled = true;
    previewValuesButton.disabled = true;

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(fillDataMessage, 'Could not get active tab.', 'error', false);
            return;
        }

        const dataBatch = []; // This will hold all field-value pairs to send to content.js

        // Iterate through each data row from the spreadsheet
        actualDataRows.forEach((spreadsheetRow, rowIndex) => {
            // For each mapped field type (contextKey), try to find and prepare data for its instances
            for (const contextKey in selectedMappings) {
                const mappedColumnHeader = selectedMappings[contextKey];
                const columnIndex = headers.indexOf(mappedColumnHeader);

                if (columnIndex !== -1) {
                    // Get all field instances for this group
                    const fieldsInThisGroup = groupedFormFields[contextKey];
                    // Assuming sequential filling for now: target the Nth instance for the Nth spreadsheet row
                    const targetField = fieldsInThisGroup[rowIndex]; 

                    if (targetField) {
                        const value = spreadsheetRow[columnIndex];
                        // Ensure value is a string for form filling
                        dataBatch.push({
                            id: targetField.id,
                            value: (value !== undefined && value !== null) ? String(value) : ''
                        });
                    } else {
                        console.warn(`No form field instance found for group "${contextKey}" at form row index ${rowIndex}. This spreadsheet row might not have a corresponding form field instance.`);
                    }
                } else {
                    console.warn(`Mapped column "${mappedColumnHeader}" not found in headers for group "${contextKey}". Skipping.`);
                }
            }
        });

        if (dataBatch.length === 0) {
            displayMessage(fillDataMessage, '<i class="fas fa-info-circle mr-2"></i>No data prepared for filling based on current mappings and spreadsheet data.', 'info', true);
            return;
        }

        displayMessage(fillDataMessage, `<i class="fas fa-paper-plane mr-2"></i>Sending ${dataBatch.length} fields for filling...`, 'info', true);

        // Send the entire batch to the content script
        const response = await chrome.tabs.sendMessage(tab.id, {
            action: 'fillBatch',
            dataBatch: dataBatch, // Send the array of {id, value} pairs
            fillEmptyOnly: fillEmptyOnly
        });

        if (response && response.status === 'success') {
            displayMessage(fillDataMessage, `<i class="fas fa-check-circle mr-2"></i>Data filling complete! ${response.filledCount} fields filled, ${response.skippedCount} fields skipped.`, 'success', true);
            saveLearnedMappings(); // Save successful mappings
        } else {
            console.error(`Error filling data: ${response?.message || 'Unknown error.'}`);
            displayMessage(fillDataMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error filling data: ${response?.message || 'Unknown error.'}`, "error", true);
        }

    } catch (error) {
        console.error("Error filling data:", error);
        displayMessage(fillDataMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error filling data: ${error.message}.`, 'error', true);
    } finally {
        fillDataButton.disabled = false;
        fillDataButton.innerHTML = '<i class="fas fa-paper-plane mr-2"></i>Fill Data';
        oneClickAutofillButton.disabled = false;
        oneClickAutofillButton.innerHTML = '<i class="fas fa-magic mr-2"></i>One-Click Autofill (Auto-Map & Fill)';
        testFillButton.disabled = false;
        previewValuesButton.disabled = false;
    }
}

/**
 * Fills only the first row of data for testing purposes.
 */
async function testFillFirstRow() {
    if (!workbook || sheetData.length <= 1) {
        displayMessage(testFillMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please upload a spreadsheet file with data first.', 'error', true);
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(testFillMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please scan for fields on the current tab first.', 'error', true);
        return;
    }

    const firstDataRow = sheetData[1]; // Get the first data row (index 1 after headers)
    if (!firstDataRow) {
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No data rows found in the spreadsheet.', 'error', true);
        return;
    }

    const fillEmptyOnly = fillEmptyOnlyCheckbox.checked;
    const dataToFillForFirstRow = [];

    const selectedMappings = {};
    mappingContainer.querySelectorAll('.group-checkbox:checked').forEach(checkbox => {
        const contextKey = checkbox.dataset.contextKey;
        const selectElement = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
        const mappedColumnHeader = selectElement ? selectElement.value : '';
        if (mappedColumnHeader) {
            selectedMappings[contextKey] = mappedColumnHeader;
        }
    });

    if (Object.keys(selectedMappings).length === 0) {
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No fields selected for filling or no column mapped.', 'error', true);
        return;
    }

    // Prepare data for the first form "row" based on the first spreadsheet data row
    for (const contextKey in selectedMappings) {
        const mappedColumnHeader = selectedMappings[contextKey];
        const columnIndex = headers.indexOf(mappedColumnHeader);

        if (columnIndex !== -1) {
            const fieldsInThisGroup = groupedFormFields[contextKey];
            const targetField = fieldsInThisGroup[0]; // Target the first instance of this grouped field
            if (targetField) {
                const value = firstDataRow[columnIndex];
                dataToFillForFirstRow.push({
                    id: targetField.id,
                    value: (value !== undefined && value !== null) ? String(value) : ''
                });
            } else {
                console.warn(`No form field instance found for group "${contextKey}" at form row index 0 for test fill.`);
            }
        }
    }

    if (dataToFillForFirstRow.length === 0) {
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No data prepared for test filling based on current mappings and first spreadsheet row.', 'info', true);
        return;
    }

    displayMessage(testFillMessage, '<i class="fas fa-spinner fa-spin mr-2"></i>Performing test fill for the first row...', 'info', true);
    testFillButton.disabled = true;
    testFillButton.innerHTML = 'Testing... <span class="loading-spinner"></span>';

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(testFillMessage, 'Could not get active tab.', 'error', false);
            return;
        }

        const response = await chrome.tabs.sendMessage(tab.id, {
            action: 'fillBatch',
            dataBatch: dataToFillForFirstRow,
            fillEmptyOnly: fillEmptyOnly
        });

        if (response && response.status === 'success') {
            displayMessage(testFillMessage, `<i class="fas fa-check-circle mr-2"></i>Test fill complete! ${response.filledCount} fields filled, ${response.skippedCount} fields skipped for the first row.`, 'success', true);
        } else {
            console.error(`Error during test fill: ${response?.message || 'Unknown error.'}`);
            displayMessage(testFillMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error during test fill: ${response?.message || 'Unknown error.'}`, "error", true);
        }

    } catch (error) {
        console.error("Error during test fill:", error);
        displayMessage(testFillMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error during test fill: ${error.message}.`, 'error', true);
    } finally {
        testFillButton.disabled = false;
        testFillButton.innerHTML = '<i class="fas fa-vial mr-2"></i>Test Fill First Row';
    }
}

/**
 * Displays a preview of mapped values for the first data row.
 */
function previewMappedValues() {
    if (!workbook || sheetData.length <= 1) {
        displayMessage(previewValuesMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please upload a spreadsheet file with data first.', 'error', true);
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(previewValuesMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Please scan for fields on the current tab first.', 'error', true);
        return;
    }

    const firstDataRow = sheetData[1];
    if (!firstDataRow) {
        displayMessage(previewValuesMessage, '<i class="fas fa-info-circle mr-2"></i>No data rows found in the spreadsheet for preview.', 'info', true);
        return;
    }

    const selectedMappings = {};
    mappingContainer.querySelectorAll('.group-checkbox:checked').forEach(checkbox => {
        const contextKey = checkbox.dataset.contextKey;
        const selectElement = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
        const mappedColumnHeader = selectElement ? selectElement.value : '';
        if (mappedColumnHeader) {
            selectedMappings[contextKey] = mappedColumnHeader;
        }
    });

    if (Object.keys(selectedMappings).length === 0) {
        displayMessage(previewValuesMessage, '<i class="fas fa-info-circle mr-2"></i>No fields selected for preview or no column mapped.', 'info', true);
        return;
    }

    let previewHtml = '<p class="font-semibold mb-2">Preview for First Data Row:</p>';
    previewHtml += '<ul class="list-disc list-inside text-left">';

    let hasPreviewData = false;
    for (const contextKey in selectedMappings) {
        const mappedColumnHeader = selectedMappings[contextKey];
        const columnIndex = headers.indexOf(mappedColumnHeader);

        if (columnIndex !== -1) {
            const fieldsInThisGroup = groupedFormFields[contextKey];
            const targetField = fieldsInThisGroup[0]; // Preview for the first instance of this grouped field
            if (targetField) {
                const value = firstDataRow[columnIndex];
                let displayName = targetField.labelText || targetField.name || targetField.placeholder || 'Unnamed Field';
                displayName = displayName.replace(/\[\d+\]/g, '').trim();
                if (displayName.startsWith('stable-id-')) displayName = 'Unnamed Field Instance';

                previewHtml += `<li><strong>${displayName}</strong> (mapped to "${mappedColumnHeader}"): <code>${(value !== undefined && value !== null) ? String(value) : '[Empty]'}</code></li>`;
                hasPreviewData = true;
            }
        }
    }
    previewHtml += '</ul>';

    if (hasPreviewData) {
        displayMessage(previewValuesMessage, previewHtml, 'info', true); // Pass true for raw HTML
    } else {
        displayMessage(previewValuesMessage, '<i class="fas fa-info-circle mr-2"></i>No preview data available based on current selections.', 'info', true);
    }
}


/**
 * Displays a temporary message in a designated message box.
 * @param {HTMLElement} element - The message box element.
 * @param {string} message - The message to display.
 * @param {'success'|'error'|'info'} type - The type of message for styling.
 * @param {boolean} isHtml - If true, message is treated as HTML.
 */
function displayMessage(element, message, type, isHtml = false) {
    if (isHtml) {
        element.innerHTML = message;
    } else {
        element.textContent = message;
    }
    element.className = `message-box mt-3 ${type === 'success' ? 'message-success' : type === 'error' ? 'message-error' : 'message-info'}`;
    element.classList.remove('hidden');
    // Ensure dark mode class is applied if body is dark
    if (document.body.classList.contains('dark')) {
        element.classList.add('dark');
    } else {
        element.classList.remove('dark');
    }

    // For preview message, keep it longer, for others, 5 seconds
    const duration = element.id === 'previewValuesMessage' ? 10000 : 5000;
    setTimeout(() => {
        element.classList.add('hidden');
        element.innerHTML = ''; // Clear content when hidden
    }, duration);
}

/**
 * Displays a temporary message in the specific fileStatusMessage element.
 * @param {string} message - The message to display.
 * @param {'success'|'error'|'info'} type - The type of message for styling.
 * @param {HTMLElement} element - The specific element to display the message (e.g., fileStatusMessage)
 * @param {boolean} isHtml - If true, message is treated as HTML.
 */
function displayFileStatusMessage(message, type, element, isHtml = false) {
    if (isHtml) {
        fileMessage.innerHTML = message;
    } else {
        fileMessage.textContent = message;
    }

    element.classList.remove('hidden');
    element.className = 'file-status mt-4'; // Reset class for styling
    element.classList.add(`message-${type}`);

    // Set icon based on type
    if (type === 'success') {
        fileStatusIcon.className = 'fas fa-check-circle text-xl mr-3';
    } else if (type === 'error') {
        fileStatusIcon.className = 'fas fa-exclamation-circle text-xl mr-3';
    } else if (type === 'info') {
        fileStatusIcon.className = 'fas fa-spinner fa-spin text-xl mr-3';
    }

    // Apply dark mode class if body is dark
    if (document.body.classList.contains('dark')) {
        element.classList.add('dark');
    } else {
        element.classList.remove('dark');
    }

    setTimeout(() => {
        element.classList.add('hidden');
        fileMessage.innerHTML = ''; // Clear content when hidden
        fileStatusIcon.className = ''; // Clear icon
    }, 5000); // Hide after 5 seconds
}
