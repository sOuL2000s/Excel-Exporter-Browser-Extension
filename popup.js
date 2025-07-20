// Global variables to store parsed data and headers
let workbook = null;
let sheetData = []; // Stores all data from the first sheet (including headers as first row)
let headers = [];   // Stores the first row (headers)
let availableFormFields = []; // Stores fields found on the active tab
let groupedFormFields = {}; // Stores fields grouped by their 'surroundingText' or derived name
let learnedMappings = {}; // Stores user's preferred mappings for schema learning

// DOM Elements
const fileInput = document.getElementById('fileInput');
const fileNameDisplay = document.getElementById('fileNameDisplay');
const fileUploadMessage = document.getElementById('fileUploadMessage');
const dataDisplaySection = document.getElementById('dataDisplaySection');
const headersDisplay = document.getElementById('headersDisplay');
const scanFieldsButton = document.getElementById('scanFieldsButton');
const scanMessage = document.getElementById('scanMessage');
const fieldMappingSection = document.getElementById('fieldMappingSection');
const mappingContainer = document.getElementById('mappingContainer');
const fillDataButton = document.getElementById('fillDataButton');
const fillDataMessage = document.getElementById('fillDataMessage');
const fillEmptyOnlyCheckbox = document.getElementById('fillEmptyOnlyCheckbox');
const oneClickAutofillButton = document.getElementById('oneClickAutofillButton'); // Autofill button
const testFillButton = document.getElementById('testFillButton'); // NEW: Test Fill button
const previewValuesButton = document.getElementById('previewValuesButton'); // NEW: Preview button
const testFillMessage = document.getElementById('testFillMessage'); // NEW: Test Fill message box
const previewValuesMessage = document.getElementById('previewValuesMessage'); // NEW: Preview message box

// --- Event Listeners ---

// File input change
fileInput.addEventListener('change', handleFile);

// Drag and drop functionality for file input
fileInput.parentElement.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.currentTarget.classList.add('border-blue-500', 'text-blue-600');
});
fileInput.parentElement.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('border-blue-500', 'text-blue-600');
});
fileInput.parentElement.addEventListener('drop', (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('border-blue-500', 'text-blue-600');
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
fillDataButton.addEventListener('click', () => fillDataInTab(false)); // Manual fill

// One-Click Autofill button click
oneClickAutofillButton.addEventListener('click', () => fillDataInTab(true)); // Auto-fill

// NEW: Test Fill button click
testFillButton.addEventListener('click', () => testFillFirstRow());

// NEW: Preview Values button click
previewValuesButton.addEventListener('click', () => previewMappedValues());

// Load learned mappings on startup
document.addEventListener('DOMContentLoaded', loadLearnedMappings);

// --- Functions ---

/**
 * Handles the file selection and reads its content.
 */
function handleFile() {
    const file = fileInput.files[0];
    if (!file) {
        displayMessage(fileUploadMessage, 'No file selected.', 'error');
        return;
    }

    fileNameDisplay.textContent = file.name;
    displayMessage(fileUploadMessage, `Reading "${file.name}"...`, 'info');

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
                displayMessage(fileUploadMessage, 'The selected file is empty or could not be parsed.', 'error');
                dataDisplaySection.classList.add('hidden');
                oneClickAutofillButton.classList.add('hidden'); // Hide autofill if no data
                return;
            }

            // The first row is the headers
            headers = sheetData[0];
            
            displayHeaders(headers);
            dataDisplaySection.classList.remove('hidden');
            displayMessage(fileUploadMessage, `File "${file.name}" loaded successfully.`, 'success');

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
            displayMessage(fileUploadMessage, `Error reading file: ${error.message}. Please ensure it's a valid spreadsheet format.`, 'error');
            dataDisplaySection.classList.add('hidden');
            oneClickAutofillButton.classList.add('hidden');
        }
    };

    reader.onerror = function(e) {
        console.error("FileReader error:", e);
        displayMessage(fileUploadMessage, `Error reading file: ${e.target.error.name}.`, 'error');
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
            span.className = 'inline-block bg-blue-100 text-blue-800 text-xs font-medium mr-2 px-2.5 py-0.5 rounded-full mb-2';
            span.textContent = header;
            headersDisplay.appendChild(span);
        });
    } else {
        headersDisplay.textContent = 'No headers found in the first row.';
    }
}

/**
 * Sends a message to the content script to scan for form fields.
 */
async function scanCurrentTabFields() {
    if (!headers || headers.length === 0) {
        displayMessage(scanMessage, 'Please upload a file with headers first.', 'error');
        return;
    }

    displayMessage(scanMessage, 'Scanning current tab for fields...', 'info');
    scanFieldsButton.disabled = true; // Disable button during scan
    scanFieldsButton.innerHTML = 'Scanning... <span class="loading-spinner"></span>';
    oneClickAutofillButton.classList.add('hidden'); // Hide autofill button during scan

    try {
        // Get the active tab
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(scanMessage, 'Could not get active tab.', 'error');
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
                groupedFormFields = groupFieldsBySignature(availableFormFields); // NEW: Use smarter grouping
                setupFieldMapping(groupedFormFields, headers); // Call existing setup to build UI with grouped fields
                autoMapFields(groupedFormFields, headers); // Call auto-map to pre-select dropdowns for groups
                fieldMappingSection.classList.remove('hidden');
                displayMessage(scanMessage, `Found ${availableFormFields.length} fields on the page, grouped into ${Object.keys(groupedFormFields).length} sections. Attempting auto-mapping.`, 'success');
                oneClickAutofillButton.classList.remove('hidden'); // Show autofill button if fields found
            } else {
                fieldMappingSection.classList.add('hidden');
                displayMessage(scanMessage, 'No input fields found on the current tab.', 'info');
                oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if no fields
            }
        } else {
            fieldMappingSection.classList.add('hidden');
            displayMessage(scanMessage, 'Failed to get fields from the current tab.', 'error');
            oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if error
        }
    } catch (error) {
        console.error("Error scanning fields:", error);
        displayMessage(scanMessage, `Error scanning fields: ${error.message}.`, 'error');
        fieldMappingSection.classList.add('hidden');
        oneClickAutofillButton.classList.add('hidden'); // Hide autofill button if error
    } finally {
        scanFieldsButton.disabled = false;
        scanFieldsButton.innerHTML = 'Scan Current Tab for Fields';
    }
}

/**
 * NEW: Generates a comprehensive signature for a form field using multiple attributes.
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
        field.autocomplete, // Include autocomplete for better context
        field.surroundingText // This is already quite good, but combine with others
    ].filter(Boolean) // Remove empty strings
     .map(str => str.toLowerCase().replace(/[^a-z0-9\s]/g, '').trim()) // Clean and normalize
     .filter(str => str.length > 1); // Remove very short, non-descriptive parts

    // Use a Set to ensure unique parts, then join
    return [...new Set(parts)].join(" | ");
}

/**
 * NEW: Groups form fields by their generated signature to create logical sections.
 * @param {Array<Object>} formFields - Array of field objects from content.js.
 * @returns {Object} An object where keys are grouping contexts (signatures) and values are arrays of fields.
 */
function groupFieldsBySignature(formFields) {
    const groups = {};
    formFields.forEach(field => {
        const signature = generateFieldSignature(field);
        // Use a fallback if the signature is empty or too generic
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
        return;
    }

    groupKeys.forEach(contextKey => {
        const fieldsInGroup = groupedFields[contextKey];
        const mappingGroupItem = document.createElement('div');
        mappingGroupItem.className = 'field-mapping-group-item mb-6 p-4 border border-gray-200 rounded-lg bg-white'; // Styling for groups

        const groupHeader = document.createElement('h3');
        groupHeader.className = 'text-md font-semibold text-gray-800 mb-3 flex items-center'; // Added flex for badge
        groupHeader.textContent = contextKey; // The grouping context (e.g., "Personal Information" or "Description")
        mappingGroupItem.appendChild(groupHeader);

        const groupControl = document.createElement('div');
        groupControl.className = 'flex items-center space-x-2 mb-3';

        const label = document.createElement('label');
        label.className = 'checkbox-label flex-shrink-0';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'group-checkbox mr-2';
        checkbox.dataset.contextKey = contextKey; // Store the context key
        checkbox.checked = false; // Initially unchecked

        const span = document.createElement('span');
        span.className = 'text-sm text-gray-700';
        span.textContent = `Map fields for "${contextKey}"`;

        label.appendChild(checkbox);
        label.appendChild(span);
        groupControl.appendChild(label);

        const select = document.createElement('select');
        select.className = 'group-mapper ml-auto flex-grow';
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

        // Display individual fields within this group (read-only)
        const fieldListContainer = document.createElement('div');
        fieldListContainer.className = 'mt-2 text-xs text-gray-600';
        fieldListContainer.innerHTML = '<p class="font-medium mb-1">Instances of this field on page:</p>';
        fieldsInGroup.forEach(field => {
            // Determine a user-friendly name for display, avoiding generated IDs
            let displayName = field.labelText || field.name || field.placeholder || field.ariaLabel || field.title || 'Unnamed Field Instance';
            // Clean up name attribute if it contains array-like indexing
            displayName = displayName.replace(/\[\d+\]/g, '').trim();

            if (displayName.startsWith('stable-id-') || displayName.startsWith('generated-id-')) {
                displayName = 'Unnamed Field Instance'; // Hide internal IDs from user
            }

            const fieldSpan = document.createElement('span');
            fieldSpan.className = 'inline-block bg-gray-100 text-gray-700 px-2 py-0.5 rounded-full mr-1 mb-1';
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
                currentSelect.value = ''; // Clear selection if unchecked
            }
        });
    });
}

/**
 * NEW: Performs intelligent auto-mapping between grouped form fields and spreadsheet headers using Fuse.js.
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
            const groupHeaderElement = checkbox.closest('.field-mapping-group-item').querySelector('h3'); // Get the h3 element for the badge

            if (!checkbox || !select) return;

            // 1. Try to apply learned mapping first
            if (learnedMappings[contextKey] && headersArray.includes(learnedMappings[contextKey])) {
                checkbox.checked = true;
                select.disabled = false;
                select.value = learnedMappings[contextKey];
                addAutoMappedBadge(groupHeaderElement, 'Learned');
                autoMappedCount++;
                console.log(`Auto-mapping (Learned): "${contextKey}" -> "${learnedMappings[contextKey]}"`);
                return; // Skip fuzzy matching if learned mapping applied
            }

            // 2. Fallback to fuzzy matching
            const result = fuse.search(contextKey)[0];
            if (result && result.score < 0.4) { // Confidence threshold
                const bestMatchHeader = result.item.header;
                checkbox.checked = true;
                select.disabled = false;
                select.value = bestMatchHeader;
                addAutoMappedBadge(groupHeaderElement, `Score: ${result.score.toFixed(2)}`);
                autoMappedCount++;
                console.log(`Auto-mapping (Fuzzy): "${contextKey}" -> "${bestMatchHeader}" (Score: ${result.score.toFixed(2)})`);
            } else {
                // If not auto-mapped, ensure it's unchecked and disabled
                checkbox.checked = false;
                select.disabled = true;
                select.value = '';
                removeAutoMappedBadge(groupHeaderElement);
                console.log(`No strong auto-mapping for "${contextKey}" (Best score: ${result?.score.toFixed(2) || 'N/A'})`);
            }
        });
        displayMessage(scanMessage, `Auto-mapping complete. ${autoMappedCount} fields auto-mapped. Review and adjust if needed.`, 'success');
    }).catch(error => {
        console.error("Error loading learned mappings for auto-map:", error);
        displayMessage(scanMessage, `Error loading learned mappings: ${error.message}. Auto-mapping proceeded without them.`, 'error');
    });
}

/**
 * NEW: Adds an "Auto-Matched" badge to the group header.
 * @param {HTMLElement} groupHeaderElement - The H3 element of the group.
 * @param {string} text - Text to display in the badge (e.g., "Learned", "Score: 0.15").
 */
function addAutoMappedBadge(groupHeaderElement, text) {
    let badge = groupHeaderElement.querySelector('.auto-matched-badge');
    if (!badge) {
        badge = document.createElement('span');
        badge.className = 'auto-matched-badge';
        groupHeaderElement.appendChild(badge);
    }
    badge.textContent = `Auto-Matched (${text})`;
}

/**
 * NEW: Removes the "Auto-Matched" badge from the group header.
 * @param {HTMLElement} groupHeaderElement - The H3 element of the group.
 */
function removeAutoMappedBadge(groupHeaderElement) {
    const badge = groupHeaderElement.querySelector('.auto-matched-badge');
    if (badge) {
        badge.remove();
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
        if (selectElement && selectElement.value) {
            currentMappings[contextKey] = selectElement.value;
        }
    });

    try {
        await chrome.storage.sync.set({ learnedMappings: currentMappings });
        console.log('Learned mappings saved:', currentMappings);
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
        displayMessage(fillDataMessage, 'Please upload a spreadsheet file first.', 'error');
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(fillDataMessage, 'Please scan for fields on the current tab first.', 'error');
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
        displayMessage(fillDataMessage, 'No fields selected for filling or no column mapped to selected groups.', 'error');
        return;
    }

    displayMessage(fillDataMessage, 'Preparing data for filling...', 'info');
    fillDataButton.disabled = true;
    oneClickAutofillButton.disabled = true;
    testFillButton.disabled = true;
    previewValuesButton.disabled = true;

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(fillDataMessage, 'Could not get active tab.', 'error');
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
            displayMessage(fillDataMessage, 'No data prepared for filling based on current mappings and spreadsheet data.', 'info');
            return;
        }

        displayMessage(fillDataMessage, `Sending ${dataBatch.length} fields for filling...`, 'info');

        // Send the entire batch to the content script
        const response = await chrome.tabs.sendMessage(tab.id, {
            action: 'fillBatch',
            dataBatch: dataBatch, // Send the array of {id, value} pairs
            fillEmptyOnly: fillEmptyOnly
        });

        if (response && response.status === 'success') {
            displayMessage(fillDataMessage, `Data filling complete! ${response.filledCount} fields filled, ${response.skippedCount} fields skipped.`, 'success');
            saveLearnedMappings(); // Save successful mappings
        } else {
            console.error(`Error filling data: ${response?.message || 'Unknown error.'}`);
            displayMessage(fillDataMessage, `Error filling data: ${response?.message || 'Unknown error.'}`, 'error');
        }

    } catch (error) {
        console.error("Error filling data:", error);
        displayMessage(fillDataMessage, `Error filling data: ${error.message}.`, 'error');
    } finally {
        fillDataButton.disabled = false;
        oneClickAutofillButton.disabled = false;
        testFillButton.disabled = false;
        previewValuesButton.disabled = false;
    }
}

/**
 * NEW: Fills only the first row of data for testing purposes.
 */
async function testFillFirstRow() {
    if (!workbook || sheetData.length <= 1) {
        displayMessage(testFillMessage, 'Please upload a spreadsheet file with data first.', 'error');
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(testFillMessage, 'Please scan for fields on the current tab first.', 'error');
        return;
    }

    const firstDataRow = sheetData[1]; // Get the first data row (index 1 after headers)
    if (!firstDataRow) {
        displayMessage(testFillMessage, 'No data rows found in the spreadsheet.', 'error');
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
        displayMessage(testFillMessage, 'No fields selected for filling or no column mapped.', 'error');
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
        displayMessage(testFillMessage, 'No data prepared for test filling based on current mappings and first spreadsheet row.', 'info');
        return;
    }

    displayMessage(testFillMessage, 'Performing test fill for the first row...', 'info');
    testFillButton.disabled = true;

    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) {
            displayMessage(testFillMessage, 'Could not get active tab.', 'error');
            return;
        }

        const response = await chrome.tabs.sendMessage(tab.id, {
            action: 'fillBatch',
            dataBatch: dataToFillForFirstRow,
            fillEmptyOnly: fillEmptyOnly
        });

        if (response && response.status === 'success') {
            displayMessage(testFillMessage, `Test fill complete! ${response.filledCount} fields filled, ${response.skippedCount} fields skipped for the first row.`, 'success');
        } else {
            console.error(`Error during test fill: ${response?.message || 'Unknown error.'}`);
            displayMessage(testFillMessage, `Error during test fill: ${response?.message || 'Unknown error.'}`, 'error');
        }

    } catch (error) {
        console.error("Error during test fill:", error);
        displayMessage(testFillMessage, `Error during test fill: ${error.message}.`, 'error');
    } finally {
        testFillButton.disabled = false;
    }
}

/**
 * NEW: Displays a preview of mapped values for the first data row.
 */
function previewMappedValues() {
    if (!workbook || sheetData.length <= 1) {
        displayMessage(previewValuesMessage, 'Please upload a spreadsheet file with data first.', 'error');
        return;
    }
    if (Object.keys(groupedFormFields).length === 0) {
        displayMessage(previewValuesMessage, 'Please scan for fields on the current tab first.', 'error');
        return;
    }

    const firstDataRow = sheetData[1];
    if (!firstDataRow) {
        displayMessage(previewValuesMessage, 'No data rows found in the spreadsheet for preview.', 'error');
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
        displayMessage(previewValuesMessage, 'No fields selected for preview or no column mapped.', 'error');
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
        displayMessage(previewValuesMessage, 'No preview data available based on current selections.', 'info');
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
    setTimeout(() => {
        element.classList.add('hidden');
        element.innerHTML = ''; // Clear content when hidden
    }, 10000); // Hide after 10 seconds for preview, 5 for others
}

// Initial state: hide data display section and mapping section
dataDisplaySection.classList.add('hidden');
fieldMappingSection.classList.add('hidden');
oneClickAutofillButton.classList.add('hidden'); // Ensure autofill button is hidden initially

