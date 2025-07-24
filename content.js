/**
 * content.js
 * This script runs in the context of the web page.
 * It listens for messages from the extension's popup and performs DOM manipulations.
 */

// Global error handler for better debugging in the content script context
window.onerror = function (msg, url, lineNo, columnNo, error) {
    console.error('Global Error (content.js):', { msg, url, lineNo, columnNo, error });
};

// Listen for messages from the popup script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'scanFields') {
        const fields = scanForFormFields();
        sendResponse({ fields: fields });
        return true; // Keep the message channel open for async response
    } else if (request.action === 'fillBatch') { // Handle batched filling
        const { dataBatch, fillEmptyOnly } = request;
        const result = fillFormFieldsBatch(dataBatch, fillEmptyOnly);
        sendResponse(result);
        return true; // Keep the message channel open for async response
    } else if (request.action === "scanClickables") { // Handle scanClickables
        const elements = getClickableElements();
        sendResponse({ clickables: elements });
        return true; // Keep the message channel open for async response
    } else if (request.action === "performClick") { // Handle performClick
        // Call the async function and then send the response
        clickElementByStableId(request.stableId, request.count).then(result => {
            sendResponse({
                status: result.success ? "success" : "error",
                message: result.message || ""
            });
        });
        return true; // This is crucial: indicates that sendResponse will be called asynchronously
    }
    // Return true to indicate that the response will be sent asynchronously
    // (already handled above for each specific action that might involve async operations)
});

/**
 * Scans the current web page for input fields, textareas, and select elements.
 * Gathers relevant information about each field, including enhanced metadata and a grouping context.
 * @returns {Array<Object>} An array of objects, each representing a form field.
 */
function scanForFormFields() {
    const fields = [];
    // Select all relevant form elements, excluding hidden, submit, button, and reset types.
    const elements = document.querySelectorAll('input:not([type="hidden"]):not([type="submit"]):not([type="button"]):not([type="reset"]), textarea, select');

    elements.forEach((element, index) => {
        // Generate a more stable and readable ID if the element doesn't have one.
        // Prioritize existing ID, then name, then placeholder, then a generated stable ID.
        let fieldId = element.id || element.name || element.placeholder || `stable-id-${index}-${Date.now()}`;
        // Ensure the ID is unique and valid for DOM selection
        if (!element.id) {
            // Make sure the generated ID is unique on the page if it's not already
            let tempId = fieldId;
            let counter = 0;
            while (document.getElementById(tempId)) {
                tempId = `${fieldId}-${counter++}`;
            }
            element.id = tempId;
            fieldId = tempId;
        }

        let labelText = '';

        // --- Prioritize table headers (<th>) for fields inside table cells (<td>) ---
        let currentCell = element.closest('td');
        if (currentCell) {
            const table = currentCell.closest('table');
            if (table) {
                const cellIndex = Array.from(currentCell.parentElement.children).indexOf(currentCell);
                // Try to find a corresponding <th> in the same column within the table's thead
                const headerCell = table.querySelector(`thead tr th:nth-child(${cellIndex + 1})`);
                if (headerCell && headerCell.textContent.trim().length > 0) {
                    labelText = headerCell.textContent.trim();
                }
            }
        }

        // Prioritize more direct and semantic labels if table header not found or empty
        if (!labelText) {
            if (element.getAttribute('aria-label')) {
                labelText = element.getAttribute('aria-label').trim();
            } else if (element.placeholder) {
                labelText = element.placeholder.trim();
            } else if (element.title) {
                labelText = element.title.trim();
            } else if (element.name) {
                // Clean up name attribute if it contains array-like indexing
                labelText = element.name.replace(/\[\d+\]/g, '').trim();
            }
        }

        // Fallback to finding an associated label using 'for' attribute
        if (!labelText && element.id) {
            const label = document.querySelector(`label[for="${element.id}"]`);
            if (label) {
                labelText = label.textContent.trim();
            }
        }

        // If no 'for' label, check parent elements for label text or aria-label
        if (!labelText) {
            let parent = element.parentElement;
            while (parent && parent.tagName !== 'BODY') {
                // Check for a <label> element that contains the current element
                const potentialLabel = parent.querySelector(`label:has(#${element.id})`);
                if (potentialLabel) {
                    labelText = potentialLabel.textContent.trim();
                    break;
                }
                // Check for a label or text directly within the parent that might serve as a label
                const parentText = Array.from(parent.childNodes)
                    .filter(node => node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0)
                    .map(node => node.textContent.trim())
                    .join(' ');
                if (parentText) {
                    labelText = parentText;
                    break;
                }
                parent = parent.parentElement;
            }
        }
        
        // --- Collect more metadata for intelligent auto-mapping and grouping context ---
        let surroundingText = '';

        // Prioritize text from immediate previous sibling elements that might be headings or prominent labels
        let prevSibling = element.previousElementSibling;
        while (prevSibling && (prevSibling.tagName === 'LABEL' || prevSibling.tagName === 'P' || prevSibling.tagName === 'DIV' || prevSibling.tagName.match(/^H[1-6]$/))) {
            const text = prevSibling.textContent.trim();
            if (text.length > 0) {
                surroundingText = text + ' ' + surroundingText;
            }
            // Stop if we hit a non-text/label/heading element or a form boundary
            if (!prevSibling.previousElementSibling || prevSibling.previousElementSibling.tagName === 'FORM') {
                break;
            }
            prevSibling = prevSibling.previousElementSibling;
        }

        // Also check parent elements for prominent text/labels, especially for table headers (re-check if not already found)
        if (!surroundingText && currentCell) { // Only if we are in a table cell and surroundingText is still empty
            const table = currentCell.closest('table');
            if (table) {
                const cellIndex = Array.from(currentCell.parentElement.children).indexOf(currentCell);
                const headerCell = table.querySelector(`thead tr th:nth-child(${cellIndex + 1})`);
                if (headerCell && headerCell.textContent.trim().length > 0) {
                    surroundingText = headerCell.textContent.trim();
                }
            }
        }

        // Check for general parent text/labels if still no strong context
        let currentParent = element.parentElement;
        while (currentParent && currentParent.tagName !== 'BODY') {
            const parentLabel = currentParent.querySelector('label');
            if (parentLabel && parentLabel.textContent.trim().length > 0 && !surroundingText.includes(parentLabel.textContent.trim())) {
                surroundingText = parentLabel.textContent.trim() + ' ' + surroundingText;
            }

            const parentHeading = currentParent.querySelector('h1, h2, h3, h4, h5, h6');
            if (parentHeading && parentHeading.textContent.trim().length > 0 && !surroundingText.includes(parentHeading.textContent.trim())) {
                surroundingText = parentHeading.textContent.trim() + ' ' + surroundingText;
            }

            Array.from(currentParent.childNodes).forEach(node => {
                if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0 && !surroundingText.includes(node.textContent.trim())) {
                    surroundingText += node.textContent.trim() + ' ';
                }
            });

            currentParent = currentParent.parentElement;
        }
        surroundingText = surroundingText.trim();

        // Fallback to labelText if surroundingText is still empty
        if (!surroundingText && labelText) {
            surroundingText = labelText;
        }
        // Fallback to placeholder if still no context
        if (!surroundingText && element.placeholder) {
            surroundingText = element.placeholder.trim();
        }
        // Fallback to name if still no context
        if (!surroundingText && element.name) {
            surroundingText = element.name.trim();
        }

        fields.push({
            id: fieldId, // Internal unique ID for the element (now more stable)
            name: element.name || '',
            type: element.type || element.tagName.toLowerCase(),
            value: element.value || '', // Current value of the field
            labelText: labelText, // Best guess for a user-friendly label
            htmlId: element.id, // Explicitly store the ID
            placeholder: element.placeholder || '',
            ariaLabel: element.getAttribute('aria-label') || '',
            className: element.className || '', // All classes as a string
            dataset: element.dataset ? JSON.parse(JSON.stringify(element.dataset)) : {}, // Convert DOMStringMap to plain object
            surroundingText: surroundingText, // This will now be our primary grouping key
            autocomplete: element.getAttribute('autocomplete') || '', // HTML autocomplete attribute
            role: element.getAttribute('role') || '', // ARIA role attribute
            title: element.getAttribute('title') || '' // HTML title attribute
        });
    });
    return fields;
}

/**
 * Fills a batch of specified form fields with data.
 * @param {Array<Object>} dataBatch - An array of objects, where each object is { fieldId: valueToFill }.
 * @param {boolean} fillEmptyOnly - If true, only fill fields that are currently empty.
 * @returns {Object} An object with status, filled count, and skipped count.
 */
function fillFormFieldsBatch(dataBatch, fillEmptyOnly) {
    let filledCount = 0;
    let skippedCount = 0;

    dataBatch.forEach(fieldData => {
        const fieldId = fieldData.id; // Use 'id' as the key for the field element
        const valueToFill = fieldData.value; // Use 'value' as the value to fill

        const element = document.getElementById(fieldId);

        if (element) {
            const currentFieldValue = element.value;

            // Check if the field should be skipped based on fillEmptyOnly flag
            if (fillEmptyOnly && currentFieldValue.trim() !== '') {
                skippedCount++;
                return; // Continue to the next item in the batch
            }

            // Fill the field based on its type
            if (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA') {
                element.value = valueToFill;
                // Dispatch input/change events to trigger any framework listeners
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
                filledCount++;
            } else if (element.tagName === 'SELECT') {
                // For select elements, try to find an option with the matching value
                let optionFound = false;
                for (let i = 0; i < element.options.length; i++) {
                    // Match by value or by text content
                    if (element.options[i].value === valueToFill || element.options[i].textContent === valueToFill) {
                        element.value = element.options[i].value;
                        optionFound = true;
                        break;
                    }
                }
                if (optionFound) {
                    element.dispatchEvent(new Event('change', { bubbles: true }));
                    filledCount++;
                } else {
                    console.warn(`Select element with ID "${fieldId}" has no option for value "${valueToFill}".`);
                    skippedCount++; // Consider this skipped if value not found
                }
            } else {
                console.warn(`Unsupported field type for element with ID "${fieldId}".`);
                skippedCount++;
            }
        } else {
            console.warn(`Element with ID "${fieldId}" not found on the page.`);
            skippedCount++;
        }
    });

    return { status: 'success', filledCount: filledCount, skippedCount: skippedCount };
}

/**
 * NEW: Scans the current web page for clickable elements.
 * @returns {Array<Object>} An array of objects, each representing a clickable element.
 */
function getClickableElements() {
    // Select common clickable elements: buttons, input type="button", links with href, and elements with role="button"
    const buttons = [...document.querySelectorAll('button, input[type="button"], a[href], [role="button"], [onclick], [tabindex="0"][aria-pressed], [tabindex="0"][aria-expanded], [tabindex="0"][role]:not([role="textbox"]):not([role="searchbox"]):not([role="combobox"]):not([role="slider"])')];
    
    // Filter out hidden elements, and elements that are likely part of other controls (like text inputs within a search button container)
    const filteredButtons = buttons.filter(el => {
        const style = window.getComputedStyle(el);
        const isVisible = style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
        const rect = el.getBoundingClientRect();
        const hasSize = rect.width > 0 && rect.height > 0;
        
        // Exclude elements that are disabled or have a disabled attribute
        if (el.disabled || el.getAttribute('aria-disabled') === 'true') {
            return false;
        }

        // Exclude inputs that are not explicitly buttons
        if (el.tagName === 'INPUT' && el.type !== 'button' && el.type !== 'submit' && el.type !== 'reset') {
            return false;
        }

        // Exclude links that are just anchors with no real action (e.g., # or javascript:void(0))
        if (el.tagName === 'A' && (!el.href || el.href.trim() === '#' || el.href.startsWith('javascript:void(0)'))) {
            return false;
        }

        return isVisible && hasSize;
    });

    return filteredButtons.map((el, index) => {
        // Prefer text content, then aria-label, then title for a human-readable label
        const text = el.innerText?.trim() || "";
        const aria = el.getAttribute("aria-label")?.trim() || "";
        const title = el.getAttribute("title")?.trim() || "";
        const name = el.getAttribute("name")?.trim() || "";
        const id = el.id?.trim() || "";

        let label = text || aria || title || name || id || `Element #${index + 1}`;
        if (label.length > 50) { // Truncate long labels for display
            label = label.substring(0, 47) + '...';
        }

        // Generate a stable ID for the element if it doesn't have one or if it's not unique
        let stableId = el.id || `autoClick-${index}-${Date.now()}`;
        if (!el.id) {
            // Ensure generated ID is unique on the page
            let tempId = stableId;
            let counter = 0;
            while (document.getElementById(tempId)) {
                tempId = `${stableId}-${counter++}`;
            }
            el.id = tempId; // Assign the generated ID to the element for easier future lookup
            stableId = tempId;
        }
        
        // Add a data attribute for easier direct selection in the content script
        el.setAttribute("data-auto-click-id", stableId);

        return { text: label, stableId: stableId }; // Return the chosen label and stable ID
    });
}

/**
 * NEW: Performs a click on an element identified by its stable ID, with a delay between clicks.
 * @param {string} stableId - The stable ID of the element to click.
 * @param {number} count - The number of times to click the element.
 * @returns {Object} An object indicating success or failure and a message.
 */
async function clickElementByStableId(stableId, count) {
    // Select using the data-auto-click-id attribute
    const target = document.querySelector(`[data-auto-click-id="${stableId}"]`);
    
    if (!target) {
        return { success: false, message: "Target element not found on the page." };
    }

    try {
        for (let i = 0; i < count; i++) {
            target.click(); // Perform the click event
            await new Promise(resolve => setTimeout(resolve, 300)); // 300ms delay between clicks
        }
        return { success: true };
    } catch (error) {
        console.error("Error simulating click:", error);
        return { success: false, message: `Error during click simulation: ${error.message}` };
    }
}
