// Global variables to store parsed data and headers
let workbook = null;
let sheetData = []; // Stores all data from the first sheet (including headers as first row)
let headers = [];   // Stores the first row (headers)
let availableFormFields = []; // Stores fields found on the active tab
let groupedFormFields = {}; // Stores fields grouped by their 'surroundingText' or derived name
let learnedMappings = {}; // Stores user's preferred mappings for schema learning

// Helper to check if content script is already loaded and responsive
async function isContentScriptLoaded(tabId) {
    try {
        // Send a dummy message to content.js and expect a 'pong' response
        const response = await chrome.tabs.sendMessage(tabId, { action: "ping" });
        return response && response.status === "pong";
    } catch (e) {
        // If an error occurs (e.g., recipient disconnected, script not injected),
        // it means the content script is not loaded or not responding.
        console.warn("Content script not loaded or not responding:", e.message);
        return false;
    }
}


// DOM Elements - Tabs
const autoFillTab = document.getElementById("autoFillTab");
const autoClickTab = document.getElementById("autoClickTab");
const autoFillSection = document.getElementById("autoFillSection");
const autoClickSection = document.getElementById("autoClickSection");

// DOM Elements - Auto Fill Section
const fileInput = document.getElementById('fileInput');
const dropArea = document.getElementById('drop-area'); // Reference to the drag and drop area
const fileNameDisplay = document.getElementById('fileNameDisplay'); // Used for "Drag & Drop..." text initially, then file name
const rowCountDisplay = document.getElementById('rowCountDisplay'); // New element for row count
const fileStatusMessage = document.getElementById('fileStatusMessage'); // The combined status div
const fileMessage = document.getElementById('fileMessage'); // Span inside status div for text
const fileStatusIcon = document.getElementById('fileStatusIcon'); // Icon inside status div
const dataDisplaySection = document.getElementById('dataDisplaySection');
const headersDisplay = document.getElementById('headersDisplay');
const scanFieldsButton = document.getElementById('scanFieldsButton');
const scanMessage = document.getElementById('scanMessage');
const fieldMappingSection = document.getElementById('fieldMappingSection');
const mappingContainer = document.getElementById('mappingContainer');
const fillDataButton = document.getElementById('fillDataButton');
const fillDataMessage = document.getElementById('fillDataMessage');
const fillEmptyOnlyCheckbox = document.getElementById('fillEmptyOnlyCheckbox');
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
const selectButtonCard = document.getElementById("selectButtonCard"); // Card for step 2
const clickControlSection = document.getElementById("clickControlSection"); // Card for step 3

// DOM Elements - Theme Toggle
const themeToggle = document.getElementById('themeToggle');


// --- Event Listeners ---

// Tab switching logic
autoFillTab.addEventListener("click", () => switchTab('autoFill'));
autoClickTab.addEventListener("click", () => switchTab('autoClick'));

// Theme toggle logic
themeToggle.addEventListener('change', () => {
    // Toggle 'dark' class on the body to apply dark mode styles
    const isDarkMode = themeToggle.checked;
    document.body.classList.toggle('dark', isDarkMode);

    // Get all elements that need their 'dark' class toggled based on the theme state
    const elementsToToggle = [
        document.querySelector('.main-container'),
        ...document.querySelectorAll('.card'),
        ...document.querySelectorAll('.section-heading'),
        ...document.querySelectorAll('.card-heading'),
        ...document.querySelectorAll('.action-button'),
        ...document.querySelectorAll('.form-input'),
        ...document.querySelectorAll('.drop-area'),
        ...document.querySelectorAll('.browse-button'),
        ...document.querySelectorAll('.file-status'),
        ...document.querySelectorAll('.headers-list'),
        ...document.querySelectorAll('.headers-list span'),
        ...document.querySelectorAll('.checkbox-label'),
        ...document.querySelectorAll('.radio-label'),
        ...document.querySelectorAll('.message-box'),
        ...document.querySelectorAll('.auto-mapped-badge'),
        ...document.querySelectorAll('.tab-buttons'),
        ...document.querySelectorAll('.tab-button'),
        document.getElementById('clickableButtonsContainer'),
        ...document.querySelectorAll('#clickableButtonsContainer > div'),
        ...document.querySelectorAll('#clickableButtonsContainer label'),
        document.getElementById('clickControlSection'),
        ...document.querySelectorAll('#clickControlSection label'),
        ...document.querySelectorAll('.selected-button-highlight'),
        ...document.querySelectorAll('.slider')
    ].filter(Boolean); // Filter out nulls if elements aren't always present

    elementsToToggle.forEach(el => {
        el.classList.toggle('dark', isDarkMode);
        // Special handling for selected-button-highlight on tab switch
        if (el.classList.contains('selected-button-highlight')) {
            el.classList.toggle('dark:selected-button-highlight', isDarkMode);
        }
        // Special handling for tab buttons active state
        if (el.classList.contains('tab-button') && el.classList.contains('active')) {
            el.classList.toggle('dark:active', isDarkMode);
        }
    });

    // Store user's theme preference in local storage
    localStorage.setItem('theme', isDarkMode ? 'dark' : 'light');
});

// File input change
fileInput.addEventListener('change', handleFile);

// Drag and drop functionality for file input
dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.classList.add('border-blue-600', 'bg-blue-100', 'dark:border-indigo-500', 'dark:bg-indigo-900');
    // Add hover text color effects only when dragging over
    fileNameDisplay.classList.add('group-hover:text-blue-700', 'dark:group-hover:text-blue-300'); 
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
fillDataButton.addEventListener('click', () => fillDataInTab());

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

        // Ensure content.js is injected only once per tab
        const loaded = await isContentScriptLoaded(tab.id);
        if (!loaded) {
            await chrome.scripting.executeScript({
                target: { tabId: tab.id },
                files: ['content.js']
            });
        }

        // Now that content.js is guaranteed to be loaded, send the actual scan message
        const response = await chrome.tabs.sendMessage(tab.id, { action: "scanClickables" });

        if (response && response.clickables) {
            if (response.clickables.length > 0) {
                clickableButtonsContainer.innerHTML = ''; // Clear scanning message
                response.clickables.forEach((btn, index) => {
                    const btnDiv = document.createElement("div");
                    btnDiv.className = "flex items-center mb-1 last:mb-0 border-b border-gray-200 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-md transition-colors duration-200 p-2";
                    btnDiv.innerHTML = `
                        <input type="radio" name="clickable" value="${btn.stableId}" id="btn-${index}" class="mr-2 flex-shrink-0">
                        <label for="btn-${index}" class="text-sm text-gray-700 dark:text-gray-300 flex-grow">${btn.text}</label>
                    `;
                    clickableButtonsContainer.appendChild(btnDiv);

                    // Add event listener to highlight selected radio button
                    btnDiv.querySelector('input[type="radio"]').addEventListener('change', (e) => {
                        // Remove highlight from all other buttons
                        document.querySelectorAll('#clickableButtonsContainer > div').forEach(div => {
                            div.classList.remove('selected-button-highlight');
                            // Ensure dark mode highlight is also removed/added correctly
                            div.classList.remove('dark:selected-button-highlight'); 
                        });
                        // Add highlight to the newly selected button's parent div
                        if (e.target.checked) {
                            btnDiv.classList.add('selected-button-highlight');
                            if (document.body.classList.contains('dark')) {
                                btnDiv.classList.add('dark:selected-button-highlight');
                            }
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
            displayFileStatusMessage('<i class="fas fa-exclamation-triangle"></i>Failed to get clickable elements from the current tab. Ensure content script can run.', 'error', autoClickMessage, true);
        }
    } catch (error) {
        console.error("Error scanning buttons:", error);
        displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error scanning buttons: ${error.message}. Check console for details.`, 'error', autoClickMessage, true);
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

        // Send message to content script to perform the clicks
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
    // Load theme preference on DOMContentLoaded and apply it immediately
    const theme = localStorage.getItem('theme');
    if (theme === 'dark') {
      themeToggle.checked = true; // Set the toggle to checked state
    }
    // Manually trigger the change event to apply the theme classes on initial load
    // This ensures all dynamically added elements (like file status, mapping groups) also get the correct theme
    themeToggle.dispatchEvent(new Event('change'));
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
    // Ensure dark mode active class is applied if current theme is dark
    if (document.body.classList.contains('dark')) {
        targetTabButton.classList.add('dark:active');
    }

    // Hide all tab sections and show the selected one
    autoFillSection.classList.add('hidden');
    autoClickSection.classList.add('hidden');
    document.getElementById(`${activeTabId}Section`).classList.remove('hidden');

    // Save tab preference to chrome.storage.sync
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
        switchTab('autoFill'); // Fallback to default in case of error
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
            // Read the workbook using SheetJS
            workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet name
            const sheetName = workbook.SheetNames[0];
            // Convert the first sheet to JSON, ensuring header:1 to get raw array of arrays
            sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

            if (sheetData.length === 0) {
                displayFileStatusMessage('<i class="fas fa-exclamation-triangle"></i>The selected file is empty or could not be parsed.', 'error', fileStatusMessage, true);
                dataDisplaySection.classList.add('hidden');
                rowCountDisplay.classList.add('hidden');
                return;
            }

            // The first row is the headers
            headers = sheetData[0];
            
            // Calculate non-empty data rows
            let nonEmtpyDataRowsCount = 0;
            // Start from the second row (index 1) to skip headers
            for (let i = 1; i < sheetData.length; i++) {
                const row = sheetData[i];
                // Check if the row contains at least one non-empty cell
                // A cell is considered non-empty if its string representation after trimming is not empty
                const isRowNonEmpty = row.some(cell => String(cell).trim() !== '');
                if (isRowNonEmpty) {
                    nonEmtpyDataRowsCount++;
                }
            }

            rowCountDisplay.textContent = `Data Rows: ${nonEmtpyDataRowsCount}`;
            rowCountDisplay.classList.remove('hidden'); // Show row count

            displayHeaders(headers); // Update UI with headers
            dataDisplaySection.classList.remove('hidden');
            displayFileStatusMessage(`<i class="fas fa-check-circle"></i>File "${file.name}" loaded successfully.`, 'success', fileStatusMessage, true);

            // If fields were already scanned, re-setup mapping with new headers and re-auto-map
            if (Object.keys(groupedFormFields).length > 0) { // Check groupedFormFields for existing scan
                setupFieldMapping(groupedFormFields, headers); // Re-setup mapping with new headers
                autoMapFields(groupedFormFields, headers); // Re-run auto-map
                fieldMappingSection.classList.remove('hidden');
            }

        } catch (error) {
            console.error("Error reading file:", error);
            displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error reading file: ${error.message}. Please ensure it's a valid spreadsheet format.`, 'error', fileStatusMessage, true);
            dataDisplaySection.classList.add('hidden');
            rowCountDisplay.classList.add('hidden');
        }
    };

    reader.onerror = function(e) {
        console.error("FileReader error:", e);
        displayFileStatusMessage(`<i class="fas fa-exclamation-triangle"></i>Error reading file: ${e.target.error.name}.`, 'error', fileStatusMessage, true);
        dataDisplaySection.classList.add('hidden');
        rowCountDisplay.classList.add('hidden');
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
            // Apply theme classes based on current body theme
            span.className = `px-5 py-2 rounded-full text-base font-medium shadow-sm flex items-center transition-colors duration-200 cursor-default ${document.body.classList.contains('dark') ? 'bg-indigo-700 text-indigo-100 hover:bg-indigo-600' : 'bg-indigo-100 text-indigo-800 hover:bg-indigo-200'}`;
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

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(scanMessage, 'Could not get active tab.', 'error', false);
            return;
        }

        // Ensure content.js is injected only once per tab
        const loaded = await isContentScriptLoaded(tab.id);
        if (!loaded) {
            await chrome.scripting.executeScript({
                target: { tabId: tab.id },
                files: ['content.js']
            });
        }

        // Now that content.js is guaranteed to be loaded, send the actual scan message
        const response = await chrome.tabs.sendMessage(tab.id, { action: 'scanFields' });

        if (response && response.fields) {
            availableFormFields = response.fields;
            if (availableFormFields.length > 0) {
                groupedFormFields = groupFieldsBySignature(availableFormFields); // Use smarter grouping
                setupFieldMapping(groupedFormFields, headers); // Call existing setup to build UI with grouped fields
                autoMapFields(groupedFormFields, headers); // Call auto-map to pre-select dropdowns for groups
                fieldMappingSection.classList.remove('hidden');
                displayMessage(scanMessage, `<i class="fas fa-check-circle mr-2"></i>Found ${availableFormFields.length} fields on the page, grouped into ${Object.keys(groupedFormFields).length} sections. Attempting auto-mapping.`, 'success', true);
            } else {
                fieldMappingSection.classList.add('hidden');
                displayMessage(scanMessage, '<i class="fas fa-info-circle mr-2"></i>No input fields found on the current tab.', 'info', true);
            }
        } else {
            fieldMappingSection.classList.add('hidden');
            displayMessage(scanMessage, '<i class="fas fa-exclamation-triangle mr-2"></i>Failed to get fields from the current tab. Ensure content script can run.', 'error', true);
        }
    } catch (error) {
        console.error("Error scanning fields:", error);
        displayMessage(scanMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error scanning fields: ${error.message}.`, 'error', true);
        fieldMappingSection.classList.add('hidden');
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
    ].filter(Boolean) // Filter out empty strings/nulls/undefineds
     .map(str => str.toLowerCase().replace(/[^a-z0-9\s]/g, '').trim()) // Clean and normalize strings
     .filter(str => str.length > 1); // Only include parts longer than 1 character

    // Use a Set to ensure unique parts and then join them
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
        // Fallback to htmlId if signature is empty, otherwise a generic label
        const groupIdentifier = signature || field.htmlId || `Unnamed Field Group (${field.type})`;
        
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
    // Remove informational text if there are groups to display
    mappingContainer.classList.remove('text-gray-500', 'dark:text-gray-400', 'text-sm');

    groupKeys.forEach(contextKey => {
        const fieldsInGroup = groupedFields[contextKey];
        const mappingGroupItem = document.createElement('div');
        // Apply card styling and theme class
        mappingGroupItem.className = `field-mapping-group-item card ${document.body.classList.contains('dark') ? 'dark' : ''}`;
        
        const groupHeader = document.createElement('h3');
        groupHeader.className = `text-md font-semibold text-gray-800 dark:text-gray-200 mb-3 flex items-center card-heading ${document.body.classList.contains('dark') ? 'dark' : ''}`;
        groupHeader.textContent = contextKey;
        mappingGroupItem.appendChild(groupHeader);

        const groupControl = document.createElement('div');
        groupControl.className = 'flex flex-col sm:flex-row items-start sm:items-center space-y-2 sm:space-y-0 sm:space-x-2 mb-3';

        const label = document.createElement('label');
        // Apply theme class to label
        label.className = `checkbox-label flex-shrink-0 ${document.body.classList.contains('dark') ? 'dark' : ''}`;
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'group-checkbox mr-2';
        checkbox.dataset.contextKey = contextKey;
        checkbox.checked = false; // Initially unchecked

        const span = document.createElement('span');
        span.className = 'text-sm text-gray-700 dark:text-gray-300';
        span.textContent = `Map fields for "${contextKey}"`;

        label.appendChild(checkbox);
        label.appendChild(span);
        groupControl.appendChild(label);

        const select = document.createElement('select');
        // Apply form-input and theme class to select
        select.className = `group-mapper flex-grow mt-2 sm:mt-0 form-input ${document.body.classList.contains('dark') ? 'dark' : ''}`;
        select.dataset.contextKey = contextKey;
        select.disabled = true; // Initially disabled

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

        mappingContainer.appendChild(mappingGroupItem);

        // Event listener for group checkbox: enables/disables the select dropdown
        checkbox.addEventListener('change', (e) => {
            const currentSelect = mappingContainer.querySelector(`.group-mapper[data-context-key="${e.target.dataset.contextKey}"]`);
            currentSelect.disabled = !e.target.checked;
            if (!e.target.checked) {
                currentSelect.value = ''; // Reset selection if unchecked
                removeAutoMappedBadge(groupHeader); // Remove badge when unchecked
            } else {
                // If re-checked, try auto-mapping again for visual consistency (e.g., if user changed it manually and then re-enabled)
                autoMapFields(groupedFields, headersArray); 
            }
        });

        // Event listener for select change to save mapping and update badge
        select.addEventListener('change', () => {
            saveLearnedMappings(); // Save whenever a mapping is changed by the user
            // Update badge based on current selection, assuming it's a manual selection if changed here
            const selectedHeader = select.value;
            if (selectedHeader) {
                updateBadgeForGroup(groupHeader, 'learned', ''); // Indicate it's now a learned mapping
            } else {
                removeAutoMappedBadge(groupHeader); // Remove badge if selection is cleared
            }
        });
    });
}

/**
 * Performs intelligent auto-mapping between grouped form fields and spreadsheet headers using Fuse.js.
 * Populates the mapping dropdowns and checks the corresponding checkboxes for groups.
 * @param {Object} groupedFields - Object of grouped fields.
 * @param {string[]} headersArray - Array of header strings.
 */
async function autoMapFields(groupedFields, headersArray) {
    await loadLearnedMappings(); // Ensure learned mappings are loaded first

    // Guard against empty headers array for Fuse.js initialization
    if (!headersArray || headersArray.length === 0) {
        console.warn("Headers array is empty, cannot perform auto-mapping.");
        displayMessage(scanMessage, '<i class="fas fa-info-circle mr-2"></i>Cannot auto-map: No headers found in the uploaded file.', 'info', true);
        return;
    }

    const fuseOptions = {
        includeScore: true,
        threshold: 0.4, // Lower is stricter, 0.4 allows for some flexibility
        keys: ['header'] // Fuse will search within the 'header' property of our items
    };

    // Prepare headers for Fuse.js search
    const fuse = new Fuse(headersArray.map(h => ({ header: h })), fuseOptions);

    let autoMappedCount = 0;
    Object.keys(groupedFields).forEach(contextKey => {
        const checkbox = mappingContainer.querySelector(`.group-checkbox[data-context-key="${contextKey}"]`);
        const select = mappingContainer.querySelector(`.group-mapper[data-context-key="${contextKey}"]`);
        const groupHeaderElement = checkbox.closest('.field-mapping-group-item').querySelector('h3');

        if (!checkbox || !select) return; // Skip if elements not found

        let mappedType = 'unmapped'; // Default mapping type

        // 1. Prioritize applying a previously learned mapping
        if (learnedMappings[contextKey] && headersArray.includes(learnedMappings[contextKey])) {
            checkbox.checked = true;
            select.disabled = false;
            select.value = learnedMappings[contextKey];
            mappedType = 'learned';
            autoMappedCount++;
            console.log(`Auto-mapping (Learned): "${contextKey}" -> "${learnedMappings[contextKey]}"`);
        } else {
            // 2. Fallback to fuzzy matching if no learned mapping or learned mapping is no longer valid (e.g., header changed)
            const result = fuse.search(contextKey)[0]; // Get the best fuzzy match
            if (result && result.score < 0.4) { // Apply if confidence (score) is high enough
                const bestMatchHeader = result.item.header;
                checkbox.checked = true;
                select.disabled = false;
                select.value = bestMatchHeader;
                mappedType = 'fuzzy';
                autoMappedCount++;
                console.log(`Auto-mapping (Fuzzy): "${contextKey}" -> "${bestMatchHeader}" (Score: ${result.score.toFixed(2)})`);
            } else {
                // If no auto-mapping, ensure checkbox is unchecked and select is disabled and reset
                checkbox.checked = false;
                select.disabled = true;
                select.value = '';
                mappedType = 'unmapped';
                console.log(`No strong auto-mapping for "${contextKey}" (Best score: ${result?.score.toFixed(2) || 'N/A'})`);
            }
        }
        // Update the visual badge for the group based on the mapping type
        updateBadgeForGroup(groupHeaderElement, mappedType, mappedType === 'fuzzy' ? result.score.toFixed(2) : '');
    });
    // Display overall auto-mapping success message
    displayMessage(scanMessage, `<i class="fas fa-check-circle mr-2"></i>Auto-mapping complete. ${autoMappedCount} fields auto-mapped. Review and adjust if needed.`, 'success', true);
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
    
    // Remove all type classes first to ensure only the current one is applied
    badge.classList.remove('learned', 'fuzzy', 'unmapped');
    // Add the current type class
    badge.classList.add(type);

    // Apply dark mode class to the badge if the body is in dark mode
    if (document.body.classList.contains('dark')) {
        badge.classList.add('dark');
    } else {
        badge.classList.remove('dark');
    }

    let badgeText = '';
    if (type === 'learned') {
        badgeText = 'Auto-Matched (Learned)';
    } else if (type === 'fuzzy') {
        badgeText = `Auto-Matched (Score: ${scoreText})`;
    } else { // 'unmapped'
        badgeText = 'Unmapped';
    }
    badge.textContent = badgeText;

    // Hide badge if unmapped or no selection in the dropdown
    const selectElement = groupHeaderElement.parentElement.querySelector('.group-mapper');
    if (type === 'unmapped' || (selectElement && !selectElement.value)) {
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
        badge.classList.add('hidden'); // Simply hide it
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
            currentMappings[contextKey] = mappedColumnHeader;
        }
    });

    try {
        await chrome.storage.sync.set({ learnedMappings: currentMappings });
        console.log('Learned mappings saved:', currentMappings);
        // After saving, re-run autoMapFields to update badges based on newly learned mappings
        if (Object.keys(groupedFormFields).length > 0 && headers.length > 0) {
            autoMapFields(groupedFormFields, headers); // This will update "learned" badges
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
 * @param {boolean} isAutoFill - True if triggered by the "One-Click Autofill" button (implies auto-mapping first).
 */
async function fillDataInTab() {
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
    // Disable buttons to prevent multiple submissions
    fillDataButton.disabled = true;
    fillDataButton.innerHTML = 'Filling... <span class="loading-spinner"></span>';
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
                    // IMPORTANT: Assuming sequential filling. Target the Nth instance for the Nth spreadsheet row.
                    // If a form has multiple instances of the same field group, this fills them in order.
                    const targetField = fieldsInThisGroup[rowIndex]; 

                    if (targetField) {
                        const value = spreadsheetRow[columnIndex];
                        // Ensure value is converted to a string for form filling consistency
                        dataBatch.push({
                            id: targetField.id,
                            value: (value !== undefined && value !== null) ? String(value) : ''
                        });
                    } else {
                        // Log a warning if a corresponding form field instance isn't found for a spreadsheet row
                        console.warn(`No form field instance found for group "${contextKey}" at form row index ${rowIndex}. This spreadsheet row might not have a corresponding form field instance.`);
                    }
                } else {
                    // Log a warning if a mapped column header is not found in the spreadsheet headers
                    console.warn(`Mapped column "${mappedColumnHeader}" not found in headers for group "${contextKey}". Skipping.`);
                }
            }
        });

        if (dataBatch.length === 0) {
            displayMessage(fillDataMessage, '<i class="fas fa-info-circle mr-2"></i>No data prepared for filling based on current mappings and spreadsheet data. Check your mappings and file content.', 'info', true);
            return;
        }

        displayMessage(fillDataMessage, `<i class="fas fa-paper-plane mr-2"></i>Sending ${dataBatch.length} fields for filling...`, 'info', true);

        // Send the entire prepared batch to the content script for filling
        const response = await chrome.tabs.sendMessage(tab.id, {
            action: 'fillBatch',
            dataBatch: dataBatch, // Array of {id, value} pairs
            fillEmptyOnly: fillEmptyOnly
        });

        if (response && response.status === 'success') {
            displayMessage(fillDataMessage, `<i class="fas fa-check-circle mr-2"></i>Data filling complete! ${response.filledCount} fields filled, ${response.skippedCount} fields skipped.`, 'success', true);
            saveLearnedMappings(); // Save successful mappings for future use
        } else {
            console.error(`Error filling data: ${response?.message || 'Unknown error.'}`);
            displayMessage(fillDataMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error filling data: ${response?.message || 'Unknown error.'}`, "error", true);
        }

    } catch (error) {
        console.error("Error filling data:", error);
        displayMessage(fillDataMessage, `<i class="fas fa-exclamation-triangle mr-2"></i>Error filling data: ${error.message}.`, 'error', true);
    } finally {
        // Re-enable buttons regardless of success or failure
        fillDataButton.disabled = false;
        fillDataButton.innerHTML = '<i class="fas fa-paper-plane mr-2"></i>Fill Data';
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
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No data rows found in the spreadsheet for testing.', 'error', true);
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
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No fields selected for test filling or no column mapped.', 'error', true);
        return;
    }

    // Prepare data for the first form "row" based on the first spreadsheet data row
    for (const contextKey in selectedMappings) {
        const mappedColumnHeader = selectedMappings[contextKey];
        const columnIndex = headers.indexOf(mappedColumnHeader);

        if (columnIndex !== -1) {
            const fieldsInThisGroup = groupedFormFields[contextKey];
            const targetField = fieldsInThisGroup[0]; // Target the first instance of this grouped field for test fill
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
        displayMessage(testFillMessage, '<i class="fas fa-info-circle mr-2"></i>No data prepared for test filling based on current mappings and first spreadsheet row. Check your mappings.', 'info', true);
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
                // Using contextKey as the display name now that instances are hidden
                const displayName = contextKey; 

                previewHtml += `<li><strong>${displayName}</strong> (mapped to "${mappedColumnHeader}"): <code>${(value !== undefined && value !== null) ? String(value) : '[Empty]'}</code></li>`;
                hasPreviewData = true;
            }
        }
    }
    previewHtml += '</ul>';

    if (hasPreviewData) {
        displayMessage(previewValuesMessage, previewHtml, 'info', true); // Pass true for raw HTML
    } else {
        displayMessage(previewValuesMessage, '<i class="fas fa-info-circle mr-2"></i>No preview data available based on current selections. Select some fields and map columns.', 'info', true);
    }
}


/**
 * Displays a temporary message in a designated message box.
 * @param {HTMLElement} element - The message box element (e.g., scanMessage, fillDataMessage).
 * @param {string} message - The message content (can be HTML if isHtml is true).
 * @param {'success'|'error'|'info'} type - The type of message for styling.
 * @param {boolean} isHtml - If true, message is parsed as HTML; otherwise, as plain text.
 */
function displayMessage(element, message, type, isHtml = false) {
    if (isHtml) {
        element.innerHTML = message;
    } else {
        element.textContent = message;
    }
    // Apply base class and type-specific class
    element.className = `message-box mt-3 ${type === 'success' ? 'message-success' : type === 'error' ? 'message-error' : 'message-info'}`;
    element.classList.remove('hidden');

    // Apply dark mode class based on current body theme
    if (document.body.classList.contains('dark')) {
        element.classList.add('dark');
    } else {
        element.classList.remove('dark');
    }

    // Set a timeout to hide the message, with a longer duration for preview messages
    const duration = element.id === 'previewValuesMessage' ? 10000 : 5000;
    setTimeout(() => {
        element.classList.add('hidden');
        element.innerHTML = ''; // Clear content when hidden
    }, duration);
}

/**
 * Displays a temporary message in the specific fileStatusMessage element.
 * This is a specialized version of displayMessage for the file upload status.
 * @param {string} message - The message content (can be HTML if isHtml is true).
 * @param {'success'|'error'|'info'} type - The type of message for styling.
 * @param {HTMLElement} element - The specific element to display the message (e.g., fileStatusMessage)
 * @param {boolean} isHtml - If true, message is parsed as HTML; otherwise, as plain text.
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

    // Set icon based on message type
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

    // Hide after 5 seconds (standard duration for file status)
    setTimeout(() => {
        element.classList.add('hidden');
        fileMessage.innerHTML = ''; // Clear content when hidden
        fileStatusIcon.className = ''; // Clear icon
    }, 5000); 
}
