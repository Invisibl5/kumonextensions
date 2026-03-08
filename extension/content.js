(function() {
    'use strict';

    // Configuration
    const CONFIG = {
        clickDelay: 50, // ms between clicks
        scrollBehavior: 'smooth',
        restrictToVisible: true,
        enableLogging: true,
        mode: 'both', // 'x', 'triangle', 'both', 'reapply'
        clicksPerBox: 1, // 1 for X, 2 for triangle, 3 to clear
        reapplyFromResults: false
    };

    // State
    let isProcessing = false;
    let totalBoxes = 0;
    let clickedCount = 0;
    let currentBoxIndex = 0;

    // UI Elements
    let uiContainer = null;
    let statusText = null;
    let progressBar = null;
    let progressText = null;
    let startButton = null;
    let stopButton = null;
    let settingsPanel = null;

    // Extension on/off (full disable: no grader, no worksheet replacement)
    let extensionEnabled = true;

    // --- Functionality 2: Worksheet Setter (Break Sets) ---
    let lastToken = null;
    let lastRegisterStudySetRequest = null;
    let lastStudyResultRequest = null;
    let lastStudyResult = null;
    let lastApiResponseId = null;
    /** First GetStudyResultInfoList response per Set (key = StudentID|SubjectCD|WorksheetCD). Used for NotUpdateMax* so we send 200 not 197. */
    let frozenStudyResultBySet = {};
    let capturedStudents = [];
    let apiCallLog = [];
    let registerLog = [];
    const API_LOG_MAX = 50;
    const REGISTER_LOG_MAX = 20;
    const REGISTER_STUDY_SET_URL = 'https://instructor2.digital.kumon.com/USA/api/ATD0010P/RegisterStudySetInfo';
    /** Study plan presets: { id, name, startPage, totalPages, patternKey } */
    let studyPlanPresets = [];
    const CLIENT = {
        applicationName: 'Class-Navi',
        version: '1.0.0.0',
        programName: 'Class-Navi',
        os: typeof navigator !== 'undefined' && navigator.userAgent ? navigator.userAgent : '-',
        machineName: '-'
    };
    const PRESETS = { '10': [10], '5': [5], '5-5': [5, 5], '4-3-3': [4, 3, 3], '3-2-3-2': [3, 2, 3, 2], '2-2-2-2-2': [2, 2, 2, 2, 2], '4-4-2': [4, 4, 2], '3-3-3-1': [3, 3, 3, 1], '3-2': [3, 2], '2-2': [2, 2] };

    function syncBodyDataset() {
        if (document.body) {
            document.body.dataset.kumonExtensionEnabled = extensionEnabled ? 'true' : 'false';
        }
    }

    function injectPageScript() {
        /* inject.js is now loaded via manifest content_scripts with world: MAIN to avoid CSP blocking inline script */
    }

    try { window.__kumonBreakSetApiLog = apiCallLog; } catch (_) {}
    try { window.__kumonBreakSetRegisterLog = registerLog; } catch (_) {}
    document.addEventListener('KumonBreakSetApiCall', function(ev) {
        const d = ev.detail;
        if (!d) return;
        const entry = { ts: Date.now(), dir: d.dir, apiName: d.apiName, url: d.url, request: d.request, response: d.response, error: d.error };
        apiCallLog.push(entry);
        if (apiCallLog.length > API_LOG_MAX) apiCallLog.shift();
        const name = d.apiName || d.url || '?';
        const dir = d.dir || '';
        const isRegister = name.indexOf('RegisterStudySetInfo') !== -1;
        if (isRegister && (dir === 'FETCH_RES' || dir === 'XHR_RES') && d.response != null) {
            let req = null;
            for (let i = apiCallLog.length - 1; i >= 0; i--) {
                const e = apiCallLog[i];
                if ((e.dir === 'FETCH_REQ' || e.dir === 'XHR_REQ') && e.apiName && e.apiName.indexOf('RegisterStudySetInfo') !== -1) { req = e.request; break; }
            }
            registerLog.push({ ts: Date.now(), source: 'page', request: req, response: d.response });
            if (registerLog.length > REGISTER_LOG_MAX) registerLog.shift();
            try { document.dispatchEvent(new CustomEvent('KumonBreakSetRegisterLogUpdated')); } catch (_) {}
        }
        if (d.response && typeof d.response === 'object') {
            const res = d.response;
            const rid = res.ID != null ? res.ID : res.id;
            if (rid != null) { const n = parseInt(rid, 10); if (!isNaN(n)) lastApiResponseId = n; }
        }
    });

    // Logging utility
    function log(message, type = 'info') {
        if (CONFIG.enableLogging) {
            const prefix = `[Kumon Auto Grader]`;
            const styles = {
                info: 'color: #2196F3',
                success: 'color: #4CAF50',
                warning: 'color: #FF9800',
                error: 'color: #F44336'
            };
            console.log(`%c${prefix} ${message}`, styles[type] || styles.info);
        }
    }

    // Check current mark type in a mark box
    function getCurrentMarkType(markBoxTarget) {
        // Find the parent mark-box element
        const markBox = markBoxTarget.closest('.mark-box');
        if (!markBox) return 'none';

        // Find the mark-box-type element
        const typeElement = markBox.querySelector('.mark-box-type');
        if (!typeElement) return 'none';

        const classList = typeElement.classList;

        // Based on the HTML provided:
        // - "check" class = X mark
        // - "triangle" class = triangle mark
        // - "default" class = nothing marked
        if (classList.contains('check')) {
            return 'x'; // X mark shows as "check" class
        } else if (classList.contains('triangle')) {
            return 'triangle';
        } else if (classList.contains('default')) {
            return 'none'; // Nothing marked shows as "default" class
        } else if (classList.contains('x') || classList.contains('wrong')) {
            return 'x';
        }

        return 'none';
    }

    // Determine how many clicks needed to reach target state
    // Click cycle: none -> x (1 click) -> triangle (2 clicks) -> none (3 clicks)
    // From any state, clicking cycles: state -> next state
    function getClicksNeeded(currentState, targetState) {
        if (currentState === targetState) {
            return 0; // Already correct
        }

        // Handle check state - treat as none
        if (currentState === 'check') {
            currentState = 'none';
        }

        // Define the cycle: none -> x -> triangle -> none (repeat)
        const cycle = ['none', 'x', 'triangle'];
        const currentIndex = cycle.indexOf(currentState);
        const targetIndex = cycle.indexOf(targetState);

        if (currentIndex === -1) {
            // Unknown current state, assume we need full clicks
            return targetState === 'x' ? 1 : targetState === 'triangle' ? 2 : 0;
        }

        if (targetIndex === -1) {
            // Target is 'none' or unknown
            if (targetState === 'none') {
                // Calculate clicks to get to none (clear)
                if (currentState === 'triangle') return 1; // triangle -> none
                if (currentState === 'x') return 2; // x -> triangle -> none
                return 0; // already none
            }
            return 0;
        }

        // Calculate clicks needed
        if (targetIndex > currentIndex) {
            // Forward in cycle
            return targetIndex - currentIndex;
        } else {
            // Need to wrap around
            return (cycle.length - currentIndex) + targetIndex;
        }
    }

    // Find result boxes and parse their marks
    function findResultBoxes() {
        const selector = '.result-box';
        let resultBoxes = Array.from(document.querySelectorAll(selector));

        if (CONFIG.restrictToVisible) {
            const visibleWorksheet = document.querySelector('.worksheet-container:not([style*="display: none"])');
            if (visibleWorksheet) {
                resultBoxes = resultBoxes.filter(box => {
                    return visibleWorksheet.contains(box) || box.closest('.worksheet-container') === visibleWorksheet;
                });
            }
        }

        // Group result boxes by question (they appear in pairs)
        const resultMap = new Map();

        resultBoxes.forEach(resultBox => {
            const typeElement = resultBox.querySelector('.result-box-type');
            if (!typeElement) return;

            // Check classList for mark types - updated to match actual HTML
            const classList = typeElement.classList;
            let markType = null;

            // Check for triangle types first (most specific)
            // triangle-check means triangle over x (triangle is the last correction)
            if (classList.contains('triangle-check') ||
                classList.contains('triangle-double') ||
                classList.contains('triangle')) {
                markType = 'triangle';
            } else if (classList.contains('check')) {
                // Check marks are actually X marks in the result boxes
                markType = 'x';
            } else if (classList.contains('x') || classList.contains('wrong') ||
                      typeElement.className.includes('x') || typeElement.className.includes('wrong')) {
                markType = 'x';
            }

            if (!markType) return; // Skip if no mark type found

            // Get position to find nearest mark box
            const rect = resultBox.getBoundingClientRect();
            const centerX = rect.left + rect.width / 2;
            const centerY = rect.top + rect.height / 2;

            // Find nearest mark box
            const allMarkBoxes = Array.from(document.querySelectorAll('.mark-box-target'));
            let nearestBox = null;
            let minDistance = Infinity;

            allMarkBoxes.forEach(markBox => {
                const markRect = markBox.getBoundingClientRect();
                const markCenterX = markRect.left + markRect.width / 2;
                const markCenterY = markRect.top + markRect.height / 2;

                const distance = Math.sqrt(
                    Math.pow(centerX - markCenterX, 2) +
                    Math.pow(centerY - markCenterY, 2)
                );

                if (distance < minDistance && distance < 100) { // Within 100px
                    minDistance = distance;
                    nearestBox = markBox;
                }
            });

            if (nearestBox) {
                const boxId = nearestBox.getAttribute('id') ||
                             `${nearestBox.getBoundingClientRect().left}-${nearestBox.getBoundingClientRect().top}`;

                if (!resultMap.has(boxId)) {
                    resultMap.set(boxId, []);
                }
                resultMap.get(boxId).push({ markType, resultBox, markBox: nearestBox });
            }
        });

        // For each mark box, get the most recent (last) mark
        // Return array of {markBox, markType} pairs to preserve order and allow different marks per box
        const reapplyList = [];
        const seenBoxes = new Set();

        resultMap.forEach((marks, boxId) => {
            // Sort by DOM order (last one is most recent)
            marks.sort((a, b) => {
                const posA = a.resultBox.compareDocumentPosition(b.resultBox);
                return posA & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
            });
            const lastMark = marks[marks.length - 1];

            // Use a unique identifier to avoid duplicates
            const boxKey = `${lastMark.markBox.getBoundingClientRect().left}-${lastMark.markBox.getBoundingClientRect().top}`;
            if (!seenBoxes.has(boxKey)) {
                seenBoxes.add(boxKey);
                reapplyList.push({
                    markBox: lastMark.markBox,
                    markType: lastMark.markType,
                    position: {
                        x: lastMark.markBox.getBoundingClientRect().left,
                        y: lastMark.markBox.getBoundingClientRect().top
                    }
                });
            }
        });

        return reapplyList;
    }

    // Find all failure boxes
    function findFailureBoxes() {
        const selector = '.mark-box-target[class*="failure"]';
        let boxes = Array.from(document.querySelectorAll(selector));

        // Filter to visible worksheet if enabled
        if (CONFIG.restrictToVisible) {
            const visibleWorksheet = document.querySelector('.worksheet-container:not([style*="display: none"])');
            if (visibleWorksheet) {
                boxes = boxes.filter(box => {
                    return visibleWorksheet.contains(box) || box.closest('.worksheet-container') === visibleWorksheet;
                });
            }
        }

        // Filter to visible boxes only
        boxes = boxes.filter(box => {
            const rect = box.getBoundingClientRect();
            // Must have size and be in viewport (not at 0,0 which means off-screen)
            return rect.width > 0 && rect.height > 0 &&
                   (rect.left !== 0 || rect.top !== 0) &&
                   rect.top >= -100 && rect.left >= -100; // Allow some margin for partial visibility
        });

        return boxes;
    }

    // Find all marked boxes (boxes that have x or triangle marks)
    // For clear mode, finds ALL marked boxes across ALL pages
    function findMarkedBoxes(includeAllPages = false) {
        // Find all mark-box-target elements (not just failure ones)
        const selector = '.mark-box-target';
        let boxes = Array.from(document.querySelectorAll(selector));

        // For clear mode, don't filter by visible worksheet - get ALL pages
        if (!includeAllPages && CONFIG.restrictToVisible) {
            const visibleWorksheet = document.querySelector('.worksheet-container:not([style*="display: none"])');
            if (visibleWorksheet) {
                boxes = boxes.filter(box => {
                    return visibleWorksheet.contains(box) || box.closest('.worksheet-container') === visibleWorksheet;
                });
            }
        }

        // Filter to only boxes that have marks (x or triangle)
        boxes = boxes.filter(box => {
            const markType = getCurrentMarkType(box);
            return markType === 'x' || markType === 'triangle';
        });

        // For clear mode, include boxes even if they're off-screen (we'll scroll to them)
        if (!includeAllPages) {
            // Filter to visible boxes only
            boxes = boxes.filter(box => {
                const rect = box.getBoundingClientRect();
                // Must have size and be in viewport (not at 0,0 which means off-screen)
                return rect.width > 0 && rect.height > 0 &&
                       (rect.left !== 0 || rect.top !== 0) &&
                       rect.top >= -100 && rect.left >= -100; // Allow some margin for partial visibility
            });
        } else {
            // For clear mode, just filter out boxes with zero size (they're truly invalid)
            boxes = boxes.filter(box => {
                const rect = box.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            });
        }

        return boxes;
    }

    // Find all mark boxes (for re-finding - includes all boxes regardless of mark state)
    function findAllMarkBoxes(includeAllPages = false) {
        const selector = '.mark-box-target';
        let boxes = Array.from(document.querySelectorAll(selector));

        // For clear mode, don't filter by visible worksheet - get ALL pages
        if (!includeAllPages && CONFIG.restrictToVisible) {
            const visibleWorksheet = document.querySelector('.worksheet-container:not([style*="display: none"])');
            if (visibleWorksheet) {
                boxes = boxes.filter(box => {
                    return visibleWorksheet.contains(box) || box.closest('.worksheet-container') === visibleWorksheet;
                });
            }
        }

        // Filter to visible boxes only (unless includeAllPages is true)
        if (!includeAllPages) {
            boxes = boxes.filter(box => {
                const rect = box.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0 &&
                       (rect.left !== 0 || rect.top !== 0) &&
                       rect.top >= -100 && rect.left >= -100;
            });
        } else {
            // For clear mode, just filter out boxes with zero size
            boxes = boxes.filter(box => {
                const rect = box.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            });
        }

        return boxes;
    }

    // Click a single box - simple and direct
    async function clickBox(box, index, numClicks = 1) {
        return new Promise(async (resolve) => {
            try {
                const startTime = performance.now();

                // Scroll into view
                box.scrollIntoView({
                    behavior: CONFIG.scrollBehavior,
                    block: 'center',
                    inline: 'center'
                });

                // Wait for scroll (reduced from 200ms)
                await new Promise(r => setTimeout(r, 100));

                // Verify element is still in DOM and visible
                if (!box.isConnected) {
                    log(`Box ${index + 1}: Element not in DOM!`, 'error');
                    resolve();
                    return;
                }

                const rect = box.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0) {
                    log(`Box ${index + 1}: Element has zero size!`, 'error');
                    resolve();
                    return;
                }

                log(`Box ${index + 1}: Clicking ${numClicks} time(s) at (${rect.left.toFixed(0)}, ${rect.top.toFixed(0)})`, 'info');

                // Simple: just click the box N times with small delays
                for (let clickNum = 0; clickNum < numClicks; clickNum++) {
                    const clickStart = performance.now();

                    // Check state before click
                    const stateBefore = getCurrentMarkType(box);

                    // Native click - this is what actually works
                    box.click();

                    // Wait a bit for the click to register (reduced from 100ms)
                    await new Promise(r => setTimeout(r, 50));

                    // Check state after click
                    const stateAfter = getCurrentMarkType(box);
                    const clickTime = (performance.now() - clickStart).toFixed(0);
                    log(`Box ${index + 1} click ${clickNum + 1}/${numClicks}: ${stateBefore} → ${stateAfter} (${clickTime}ms)`, 'info');

                    // Tiny delay between clicks (reduced from 50ms)
                    if (clickNum < numClicks - 1) {
                        await new Promise(r => setTimeout(r, 30));
                    }
                }

                // Small delay after all clicks (reduced from 150ms)
                await new Promise(r => setTimeout(r, 80));

                const totalTime = (performance.now() - startTime).toFixed(0);
                log(`Box ${index + 1}: Click complete in ${totalTime}ms`, 'info');

                resolve();
            } catch (error) {
                log(`Error clicking box ${index + 1}: ${error.message}`, 'error');
                resolve();
            }
        });
    }

    // Process all boxes based on mode
    async function processAllBoxes() {
        if (isProcessing) {
            log('Already processing, please wait...', 'warning');
            return;
        }

        isProcessing = true;
        clickedCount = 0;
        currentBoxIndex = 0;

        let boxes = [];

        if (CONFIG.mode === 'reapply') {
            // Reapply mode: simple - keep making passes until everything matches
            startButton.disabled = true;
            stopButton.disabled = false;

            const maxPasses = 10;
            let pass = 0;

            while (pass < maxPasses && isProcessing) {
                pass++;
                log(`Pass ${pass}: Finding result boxes...`, 'info');

                // Get fresh list of boxes to process
                const reapplyList = findResultBoxes();

                if (reapplyList.length === 0) {
                    log('No result boxes found', 'warning');
                    break;
                }

                totalBoxes = reapplyList.length;
                updateStatus(`Pass ${pass}: Processing ${totalBoxes} boxes...`, 'processing');
                updateProgress();

                let allMatch = true;
                let processedCount = 0;

                for (let i = 0; i < reapplyList.length; i++) {
                    if (!isProcessing) break;

                    const { markType: targetState, position } = reapplyList[i];

                    // Find the mark box by position
                    const allBoxes = findFailureBoxes();
                    let markBox = null;
                    let minDistance = Infinity;

                    allBoxes.forEach(box => {
                        const rect = box.getBoundingClientRect();
                        const distance = Math.sqrt(
                            Math.pow(rect.left - position.x, 2) +
                            Math.pow(rect.top - position.y, 2)
                        );
                        if (distance < minDistance && distance < 50) {
                            minDistance = distance;
                            markBox = box;
                        }
                    });

                    if (!markBox) {
                        log(`Box ${i + 1}: Could not find mark box`, 'warning');
                        allMatch = false;
                        continue;
                    }

                    const currentState = getCurrentMarkType(markBox);

                    // Skip if already matches
                    if (currentState === targetState) {
                        log(`Box ${i + 1}: Already matches (${targetState}), skipping`, 'info');
                        processedCount++;
                        continue;
                    }

                    // Doesn't match - click it once
                    allMatch = false;
                    log(`Box ${i + 1}: Current=${currentState}, Target=${targetState}, clicking...`, 'info');

                    await clickBox(markBox, i, 1);
                    await new Promise(resolve => setTimeout(resolve, 200));

                    processedCount++;
                    currentBoxIndex = processedCount;
                    clickedCount = processedCount;
                    updateProgress();
                }

                // If everything matches, we're done!
                if (allMatch) {
                    log(`All boxes match! Completed in ${pass} passes.`, 'success');
                    updateStatus(`Completed! All boxes match (${pass} passes)`, 'success');
                    break;
                }

                // Wait before next pass
                if (pass < maxPasses) {
                    log(`Pass ${pass} complete. Some boxes still don't match. Starting pass ${pass + 1}...`, 'info');
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }

            if (pass >= maxPasses && isProcessing) {
                updateStatus(`Completed ${maxPasses} passes`, 'warning');
                log(`Completed ${maxPasses} passes. Some boxes may still not match.`, 'warning');
            }

            isProcessing = false;
            startButton.disabled = false;
            stopButton.disabled = true;
            return; // Exit early for reapply mode
        } else if (CONFIG.mode === 'clear') {
            // Clear mode: find ALL marked boxes across ALL pages
            boxes = findMarkedBoxes(true); // true = include all pages

            if (boxes.length === 0) {
                updateStatus('No marked boxes found to clear', 'warning');
                isProcessing = false;
                return;
            }

            log(`Found ${boxes.length} marked boxes to clear (across all pages)`, 'info');
            updateStatus(`Clearing ${boxes.length} marked boxes...`, 'processing');
        } else {
            // Normal mode: find failure boxes
            boxes = findFailureBoxes();

            if (boxes.length === 0) {
                updateStatus('No failure boxes found', 'warning');
                isProcessing = false;
                return;
            }

            log(`Found ${boxes.length} failure boxes`, 'info');
            updateStatus(`Processing ${boxes.length} boxes...`, 'processing');
        }

        totalBoxes = boxes.length;
        updateProgress();

        startButton.disabled = true;
        stopButton.disabled = false;

        // Process boxes sequentially with verification
        const processStartTime = performance.now();
        log(`Starting to process ${boxes.length} boxes...`, 'info');

        // Multi-pass verification: keep running passes until all boxes are correct
        let passNumber = 1;
        const maxPasses = 10; // Safety limit
        let boxesNeedingWork = [];

        while (passNumber <= maxPasses) {
            log(`\n=== PASS ${passNumber}/${maxPasses} ===`, 'info');
            boxesNeedingWork = [];

            // Re-find boxes at the start of each pass (for reapply/clear modes)
            if (CONFIG.mode === 'reapply') {
                const reapplyList = findResultBoxes();
                if (reapplyList.length === 0) {
                    log('No result boxes found for reapply mode', 'warning');
                    break;
                }
                boxes = reapplyList.map(item => item.markBox);
            } else if (CONFIG.mode === 'clear') {
                boxes = findMarkedBoxes(true); // Get all marked boxes
                if (boxes.length === 0) {
                    log('No marked boxes found - all cleared!', 'success');
                    break;
                }
            } else {
                boxes = findFailureBoxes();
            }

            log(`Pass ${passNumber}: Found ${boxes.length} boxes to process`, 'info');

            if (boxes.length === 0) {
                log(`Pass ${passNumber}: No boxes to process, done!`, 'success');
                break;
            }

            for (let i = 0; i < boxes.length; i++) {
            if (!isProcessing) {
                updateStatus('Stopped by user', 'stopped');
                break;
            }

            const boxStartTime = performance.now();
            let markBox = boxes[i]; // Use let so we can reassign when element is recreated
            const currentState = getCurrentMarkType(markBox);
            let targetState = null;
            let numClicks = 0;

            // Determine target state and clicks needed based on mode
            // (reapply mode is handled separately above)
            if (CONFIG.mode === 'x') {
                targetState = 'x';
                numClicks = getClicksNeeded(currentState, 'x');
                if (numClicks === 0) {
                    log(`Box ${i + 1}/${totalBoxes}: Already has X, skipping`, 'info');
                    clickedCount++;
                    currentBoxIndex = i + 1;
                    updateProgress();
                    continue;
                }
            } else if (CONFIG.mode === 'triangle') {
                targetState = 'triangle';
                numClicks = getClicksNeeded(currentState, 'triangle');
                if (numClicks === 0) {
                    log(`Box ${i + 1}/${totalBoxes}: Already has triangle, skipping`, 'info');
                    clickedCount++;
                    currentBoxIndex = i + 1;
                    updateProgress();
                    continue;
                }
            } else if (CONFIG.mode === 'clear') {
                targetState = 'none';
                // For clear mode, always process (even if already none, verify it)
                // We'll keep clicking until it's actually 'none'
                numClicks = getClicksNeeded(currentState, 'none');
                log(`Box ${i + 1}/${totalBoxes}: Clear mode - Current=${currentState}, Target=${targetState}, Will click until blank`, 'info');
                if (numClicks === 0 && currentState === 'none') {
                    log(`Box ${i + 1}/${totalBoxes}: Already cleared, skipping`, 'info');
                    clickedCount++;
                    currentBoxIndex = i + 1;
                    updateProgress();
                    continue;
                }
                // Force numClicks to at least 1 if not already none, so the loop runs
                if (numClicks === 0 && currentState !== 'none') {
                    numClicks = 1; // Will click at least once, loop will handle the rest
                }
            } else if (CONFIG.mode === 'both') {
                // For 'both' mode, use clicksPerBox setting
                // Determine what state we want based on clicks
                if (CONFIG.clicksPerBox === 1) {
                    targetState = 'x';
                    numClicks = getClicksNeeded(currentState, 'x');
                } else if (CONFIG.clicksPerBox === 2) {
                    targetState = 'triangle';
                    numClicks = getClicksNeeded(currentState, 'triangle');
                } else {
                    // 3 clicks = clear
                    targetState = 'none';
                    numClicks = getClicksNeeded(currentState, 'none');
                }

                if (numClicks === 0 && CONFIG.clicksPerBox !== 3) {
                    log(`Box ${i + 1}/${totalBoxes}: Already has ${targetState}, skipping`, 'info');
                    clickedCount++;
                    currentBoxIndex = i + 1;
                    updateProgress();
                   	continue;
                }
            }

            // Verify we need to click
            log(`Box ${i + 1}/${totalBoxes}: numClicks=${numClicks}, Current=${currentState}, Target=${targetState}`, 'info');
            if (numClicks > 0) {
                log(`Box ${i + 1}/${totalBoxes}: Will click ${numClicks} time(s)`, 'info');

                // Keep clicking until we reach the target state
                // IMPORTANT: Re-find the element after each click since Angular recreates DOM
                let newState = currentState;
                let totalClicks = 0;
                const maxTotalClicks = 20; // Safety limit

                // Store the initial position/identifier to re-find the element
                const initialRect = markBox.getBoundingClientRect();
                const initialX = initialRect.left;
                const initialY = initialRect.top;

                // Skip if box is off-screen (at 0,0 or negative) - UNLESS we're in clear mode
                // In clear mode, we want to clear ALL boxes, so we'll scroll to them
                if (CONFIG.mode !== 'clear' && ((initialX === 0 && initialY === 0) || initialX < -50 || initialY < -50)) {
                    log(`Box ${i + 1}: Skipping - box is off-screen at (${initialX.toFixed(0)}, ${initialY.toFixed(0)})`, 'warning');
                    clickedCount++;
                    currentBoxIndex = i + 1;
                    updateProgress();
                    continue;
                }

                // For clear mode, if box is off-screen, scroll to it first
                if (CONFIG.mode === 'clear' && ((initialX === 0 && initialY === 0) || initialX < -50 || initialY < -50)) {
                    log(`Box ${i + 1}: Box is off-screen, scrolling to it...`, 'info');
                    markBox.scrollIntoView({
                        behavior: CONFIG.scrollBehavior,
                        block: 'center',
                        inline: 'center'
                    });
                    await new Promise(resolve => setTimeout(resolve, 300)); // Wait for scroll
                    // Re-get position after scroll
                    const newRect = markBox.getBoundingClientRect();
                    const newX = newRect.left;
                    const newY = newRect.top;
                    log(`Box ${i + 1}: After scroll, position is (${newX.toFixed(0)}, ${newY.toFixed(0)})`, 'info');
                }

                log(`Box ${i + 1}: Starting at position (${initialX.toFixed(0)}, ${initialY.toFixed(0)})`, 'info');

                const loopStartTime = performance.now();

                while (newState !== targetState && totalClicks < maxTotalClicks) {
                    const attemptStartTime = performance.now();

                    // Re-find the element by position (Angular recreates DOM elements)
                    let currentBox = markBox;
                    if (!currentBox.isConnected) {
                        log(`Box ${i + 1}: Element disconnected, re-finding...`, 'warning');
                        const findStart = performance.now();
                        // Element was removed from DOM, find it again
                        // Use findAllMarkBoxes for re-finding (works for all modes)
                        // In clear mode, search all pages
                        const allBoxes = findAllMarkBoxes(CONFIG.mode === 'clear');
                        const findTime = (performance.now() - findStart).toFixed(0);
                        log(`Box ${i + 1}: Found ${allBoxes.length} boxes in ${findTime}ms`, 'info');

                        // Find box closest to original position
                        let closestBox = null;
                        let minDistance = Infinity;
                        const searchStart = performance.now();
                        allBoxes.forEach((box, idx) => {
                            const rect = box.getBoundingClientRect();
                            const distance = Math.sqrt(
                                Math.pow(rect.left - initialX, 2) +
                                Math.pow(rect.top - initialY, 2)
                            );
                            // For clear mode, use larger threshold (500px) since boxes move a lot
                            const threshold = CONFIG.mode === 'clear' ? 500 : 150;
                            if (distance < minDistance && distance < threshold) {
                                minDistance = distance;
                                closestBox = box;
                            }
                        });
                        const searchTime = (performance.now() - searchStart).toFixed(0);

                        if (closestBox) {
                            currentBox = closestBox;
                            markBox = closestBox; // Update reference
                            log(`Box ${i + 1}: ✓ Re-found at ${minDistance.toFixed(0)}px (search: ${searchTime}ms)`, 'success');
                        } else {
                            log(`Box ${i + 1}: ✗ Not found! Searched ${allBoxes.length} boxes in ${searchTime}ms, min dist: ${minDistance === Infinity ? 'N/A' : minDistance.toFixed(0) + 'px'}`, 'error');
                            // Log all box positions for debugging
                            if (CONFIG.enableLogging && allBoxes.length > 0) {
                                const positions = allBoxes.slice(0, 5).map((b, idx) => {
                                    const r = b.getBoundingClientRect();
                                    return `box${idx}:(${r.left.toFixed(0)},${r.top.toFixed(0)})`;
                                }).join(', ');
                                log(`Box ${i + 1}: First 5 box positions: ${positions}`, 'info');
                            }

                            // For clear mode, don't give up - try to find by index or use the first marked box
                            if (CONFIG.mode === 'clear' && allBoxes.length > 0) {
                                // Try to find a marked box near the original position with larger threshold
                                let fallbackBox = null;
                                let fallbackDistance = Infinity;
                                allBoxes.forEach(box => {
                                    const rect = box.getBoundingClientRect();
                                    const distance = Math.sqrt(
                                        Math.pow(rect.left - initialX, 2) +
                                        Math.pow(rect.top - initialY, 2)
                                    );
                                    const markType = getCurrentMarkType(box);
                                    // If it's still marked and closer, use it
                                    if ((markType === 'x' || markType === 'triangle') && distance < fallbackDistance && distance < 500) {
                                        fallbackDistance = distance;
                                        fallbackBox = box;
                                    }
                                });

                                if (fallbackBox) {
                                    currentBox = fallbackBox;
                                    markBox = fallbackBox;
                                    log(`Box ${i + 1}: Using fallback box at ${fallbackDistance.toFixed(0)}px (still marked)`, 'warning');
                                } else {
                                    // Last resort: use the box at the same index in the list
                                    if (i < allBoxes.length) {
                                        currentBox = allBoxes[i];
                                        markBox = allBoxes[i];
                                        log(`Box ${i + 1}: Using box at index ${i} as fallback`, 'warning');
                                    } else {
                                        log(`Box ${i + 1}: Cannot continue - no fallback available`, 'error');
                                        break;
                                    }
                                }
                            } else {
                                break;
                            }
                        }
                    }

                    // Click the current (fresh) element
                    log(`Box ${i + 1}: Click attempt ${totalClicks + 1}/${maxTotalClicks}...`, 'info');
                    await clickBox(currentBox, i, 1);
                    totalClicks++;

                    // Wait for UI to update (reduced from 200ms)
                    await new Promise(resolve => setTimeout(resolve, 100));

                    // Re-find element again and check state
                    if (!currentBox.isConnected) {
                        log(`Box ${i + 1}: Disconnected after click, re-finding...`, 'warning');
                        // Use findAllMarkBoxes for re-finding (works for all modes)
                        // In clear mode, search all pages
                        const allBoxes = findAllMarkBoxes(CONFIG.mode === 'clear');
                        let closestBox = null;
                        let minDistance = Infinity;
                        allBoxes.forEach(box => {
                            const rect = box.getBoundingClientRect();
                            const distance = Math.sqrt(
                                Math.pow(rect.left - initialX, 2) +
                                Math.pow(rect.top - initialY, 2)
                            );
                            // For clear mode, use larger threshold (500px) since boxes move a lot
                            const threshold = CONFIG.mode === 'clear' ? 500 : 150;
                            if (distance < minDistance && distance < threshold) {
                                minDistance = distance;
                                closestBox = box;
                            }
                        });
                        if (closestBox) {
                            currentBox = closestBox;
                            markBox = closestBox;
                            log(`Box ${i + 1}: Re-found after click at ${minDistance.toFixed(0)}px`, 'info');
                        } else {
                            log(`Box ${i + 1}: ✗ Lost after click!`, 'error');
                            // For clear mode, try fallback strategy
                            if (CONFIG.mode === 'clear' && allBoxes.length > 0) {
                                // Find any still-marked box near the original position
                                let fallbackBox = null;
                                let fallbackDistance = Infinity;
                                allBoxes.forEach(box => {
                                    const rect = box.getBoundingClientRect();
                                    const distance = Math.sqrt(
                                        Math.pow(rect.left - initialX, 2) +
                                        Math.pow(rect.top - initialY, 2)
                                    );
                                    const markType = getCurrentMarkType(box);
                                    if ((markType === 'x' || markType === 'triangle') && distance < fallbackDistance && distance < 500) {
                                        fallbackDistance = distance;
                                        fallbackBox = box;
                                    }
                                });

                                if (fallbackBox) {
                                    currentBox = fallbackBox;
                                    markBox = fallbackBox;
                                    log(`Box ${i + 1}: Using fallback box at ${fallbackDistance.toFixed(0)}px after click`, 'warning');
                                } else if (i < allBoxes.length) {
                                    currentBox = allBoxes[i];
                                    markBox = allBoxes[i];
                                    log(`Box ${i + 1}: Using box at index ${i} as fallback after click`, 'warning');
                                }
                            }
                        }
                    }

                    newState = getCurrentMarkType(currentBox);
                    const attemptTime = (performance.now() - attemptStartTime).toFixed(0);
                    log(`Box ${i + 1}: After click ${totalClicks}, state=${newState}, target=${targetState} (${attemptTime}ms)`, 'info');

                    // For clear mode, keep clicking until it's actually 'none'
                    // For other modes, stop when we reach target
                    if (newState === targetState) {
                        const totalTime = (performance.now() - loopStartTime).toFixed(0);
                        log(`Box ${i + 1}: ✓ Success! ${newState} in ${totalClicks} clicks (${totalTime}ms total)`, 'success');
                        break;
                    }

                    // For clear mode, if we're not at 'none' yet, keep going
                    if (CONFIG.mode === 'clear' && newState !== 'none') {
                        log(`Box ${i + 1}: Still not cleared (current: ${newState}), continuing...`, 'info');
                    }

                    // Debug: warn if state isn't changing
                    if (totalClicks > 3 && newState === currentState) {
                        log(`Box ${i + 1}: ⚠️ State stuck at ${newState} after ${totalClicks} clicks`, 'warning');
                    }
                }

                // Final check - verify the state is actually correct
                const finalState = getCurrentMarkType(markBox);
                if (finalState !== targetState) {
                    log(`Box ${i + 1}: ✗ FAILED VERIFICATION! Final state=${finalState}, expected=${targetState}`, 'error');
                    log(`Box ${i + 1}: WHY IT FAILED: Clicked ${totalClicks} times, max was ${maxTotalClicks}. State changed from ${currentState} to ${finalState}`, 'error');
                    boxesNeedingWork.push({
                        index: i,
                        box: markBox,
                        currentState: finalState,
                        targetState: targetState,
                        reason: `State is ${finalState} but should be ${targetState} after ${totalClicks} clicks`
                    });
                } else {
                    log(`Box ${i + 1}: ✓ VERIFIED! State is ${finalState} (correct)`, 'success');
                }

                // Check if we reached the target (or max clicks)
                if (newState !== targetState && totalClicks >= maxTotalClicks) {
                    log(`Box ${i + 1}: ✗ Max clicks reached. Final state=${newState}, expected=${targetState}`, 'error');
                    log(`Box ${i + 1}: WHY: Reached max clicks (${maxTotalClicks}) without reaching target state`, 'error');
                } else if (CONFIG.mode === 'clear' && newState !== 'none' && totalClicks >= maxTotalClicks) {
                    log(`Box ${i + 1}: ✗ Clear mode: Max clicks reached but still not blank. Final state=${newState}`, 'error');
                    log(`Box ${i + 1}: WHY: Clicked ${maxTotalClicks} times but box still has mark (${newState})`, 'error');
                }

                // Update count only once per box
                clickedCount++;
                currentBoxIndex = i + 1;
                updateProgress();
            } else {
                // Already correct, verify it
                const verifiedState = getCurrentMarkType(markBox);
                if (verifiedState === targetState) {
                    log(`Box ${i + 1}: ✓ Already correct (${verifiedState}), verified`, 'success');
                } else {
                    log(`Box ${i + 1}: ⚠️ Thought it was correct but verification shows ${verifiedState} (expected ${targetState})`, 'warning');
                    boxesNeedingWork.push({
                        index: i,
                        box: markBox,
                        currentState: verifiedState,
                        targetState: targetState,
                        reason: `Initial check said correct but verification shows ${verifiedState}`
                    });
                }
                clickedCount++;
                currentBoxIndex = i + 1;
                updateProgress();
            }

            const boxTime = (performance.now() - boxStartTime).toFixed(0);
            log(`Box ${i + 1}: Total time ${boxTime}ms`, 'info');

            // Reduced delay between boxes for speed
            await new Promise(resolve => setTimeout(resolve, Math.max(20, CONFIG.clickDelay / 2)));
            }

            // End of pass - check if we need another pass
            if (boxesNeedingWork.length === 0) {
                log(`\n✓ PASS ${passNumber} COMPLETE: All boxes verified correct!`, 'success');
                break;
            } else {
                log(`\n⚠️ PASS ${passNumber} COMPLETE: ${boxesNeedingWork.length} boxes still need work`, 'warning');
                log(`Boxes needing work: ${boxesNeedingWork.map(b => `Box ${b.index + 1} (${b.currentState}→${b.targetState}: ${b.reason})`).join(', ')}`, 'warning');

                if (passNumber < maxPasses) {
                    log(`\nStarting verification pass ${passNumber + 1}...`, 'info');
                    passNumber++;
                    // Reset for next pass
                    clickedCount = 0;
                    currentBoxIndex = 0;
                    await new Promise(resolve => setTimeout(resolve, 500)); // Brief pause between passes
                } else {
                    log(`\n✗ MAX PASSES REACHED: ${boxesNeedingWork.length} boxes still incorrect after ${maxPasses} passes`, 'error');
                    break;
                }
            }
        }

        if (isProcessing) {
            const totalTime = ((performance.now() - processStartTime) / 1000).toFixed(1);
            const finalBoxesNeedingWork = boxesNeedingWork.length;
            if (finalBoxesNeedingWork === 0) {
                updateStatus(`Completed! All boxes verified correct in ${passNumber} pass(es)`, 'success');
                log(`✓ COMPLETED: All boxes verified correct in ${passNumber} pass(es) in ${totalTime}s`, 'success');
            } else {
                updateStatus(`Completed with ${finalBoxesNeedingWork} boxes still incorrect`, 'warning');
                log(`⚠️ COMPLETED: ${finalBoxesNeedingWork} boxes still incorrect after ${passNumber} passes in ${totalTime}s`, 'warning');
            }
        }

        isProcessing = false;
        startButton.disabled = false;
        stopButton.disabled = true;
    }

    // Stop processing
    function stopProcessing() {
        if (isProcessing) {
            isProcessing = false;
            updateStatus('Stopping...', 'stopped');
            log('Processing stopped by user', 'warning');
        }
    }

    // Update UI status
    function updateStatus(message, type = 'info') {
        if (!statusText) return;

        statusText.textContent = message;
        statusText.className = `status-text status-${type}`;
    }

    // Update progress bar
    function updateProgress() {
        if (!progressBar || !progressText) return;

        const percentage = totalBoxes > 0 ? (currentBoxIndex / totalBoxes) * 100 : 0;
        progressBar.style.width = `${percentage}%`;
        progressText.textContent = `${currentBoxIndex} / ${totalBoxes}`;
    }

    // --- Worksheet Setter helpers ---
    function parsePattern(str) {
        if (!str || typeof str !== 'string') return [];
        const s = str.trim().replace(/[,，\s]+/g, '-');
        return s.split('-').map(x => { const n = parseInt(x, 10); return isNaN(n) || n < 1 ? 0 : n; }).filter(n => n > 0);
    }
    function expandPatternToTotal(patternSizes, totalPages) {
        const baseSum = patternSizes.reduce((a, b) => a + b, 0);
        if (baseSum <= 0) return null;
        const fullRepeats = Math.floor(totalPages / baseSum);
        if (fullRepeats < 1) return null;
        const expanded = [];
        for (let r = 0; r < fullRepeats; r++) for (let i = 0; i < patternSizes.length; i++) expanded.push(patternSizes[i]);
        return expanded;
    }
    /** Build InsertSetInfoList - multiple break items, ALL with the SAME StudyScheduleIndex. No set may cross a decade (e.g. 60-61); such chunks are split at the boundary. */
    function buildInsertSetInfoList(startPage, totalPages, patternSizes, nextStudyScheduleIndex) {
        const sum = patternSizes.reduce((a, b) => a + b, 0);
        if (sum !== totalPages) return null;
        const index = nextStudyScheduleIndex != null ? nextStudyScheduleIndex : 1;
        const list = []; let from = startPage;
        for (let i = 0; i < patternSizes.length; i++) {
            const n = patternSizes[i];
            const to = from + n - 1;
            const blockFrom = Math.floor((from - 1) / 10);
            const blockTo = Math.floor((to - 1) / 10);
            if (blockFrom === blockTo) {
                list.push({ StudyScheduleIndex: index, WorksheetNOFrom: from, WorksheetNOTo: to, GradingMethod: '1' });
            } else {
                const endOfBlock = (blockFrom + 1) * 10;
                const startOfNext = endOfBlock + 1;
                list.push({ StudyScheduleIndex: index, WorksheetNOFrom: from, WorksheetNOTo: endOfBlock, GradingMethod: '1' });
                list.push({ StudyScheduleIndex: index, WorksheetNOFrom: startOfNext, WorksheetNOTo: to, GradingMethod: '1' });
            }
            from = to + 1;
        }
        return list;
    }
    function getStudyResultData(res) { if (!res) return null; return res.data || res.Data || res; }
    function isStudyUnitCompleted(u) {
        if (!u) return false;
        if (u.StudyStatus === '6') return true;
        if (u.StudyDate || u.FinishDate) return true;
        return false;
    }
    function buildFinishTestItem(u) {
        return {
            StudyScheduleIndex: u.StudyScheduleIndex != null ? u.StudyScheduleIndex : 1,
            StudySec: u.StudySec || '1',
            WorksheetNOFrom: u.WorksheetNOFrom,
            WorksheetNOTo: u.WorksheetNOTo,
            StudyDate: u.StudyDate || null,
            StudyStatus: u.StudyStatus || '6',
            FinishDate: u.FinishDate || u.StudyDate || null,
            CompleteTime: u.CompleteTime != null ? u.CompleteTime : null,
            FirstCompleteTime: u.FirstCompleteTime != null ? u.FirstCompleteTime : u.CompleteTime,
            StandardTimeFrom: u.StandardTimeFrom != null ? u.StandardTimeFrom : null,
            StandardTimeTo: u.StandardTimeTo != null ? u.StandardTimeTo : null,
            GradingMethod: u.GradingMethod || '1',
            DownloadFlg: u.DownloadFlg != null ? u.DownloadFlg : '1',
            StudyStartTime: u.StudyStartTime || null,
            DeleteFlg: u.DeleteFlg != null ? u.DeleteFlg : '0',
            SoundFlg: u.SoundFlg != null ? u.SoundFlg : '0'
        };
    }
    /** Build RegisterStudySetInfo payload from GetStudyResultInfoList request + response. NotUpdateMaxWorksheetNO fallback = max worksheet already downloaded. */
    function buildPayloadFromStudyResult() {
        const req = typeof lastStudyResultRequest === 'string' ? (() => { try { return JSON.parse(lastStudyResultRequest); } catch (e) { return null; } })() : lastStudyResultRequest;
        if (!req || !lastStudyResult) return null;
        const res = getStudyResultData(lastStudyResult);
        if (!res) return null;
        const list = res.StudyUnitInfoList || [];
        let maxIndexWithStudyDate = 0, maxWorksheetNODownloaded = 0, maxWorksheetNOCompleted = 0;
        list.forEach(u => {
            if (u.StudyDate && u.StudyScheduleIndex != null && u.StudyScheduleIndex > maxIndexWithStudyDate) maxIndexWithStudyDate = u.StudyScheduleIndex;
            if (u.DownloadFlg === '1' && u.WorksheetNOTo != null && u.WorksheetNOTo > maxWorksheetNODownloaded) maxWorksheetNODownloaded = u.WorksheetNOTo;
            if (isStudyUnitCompleted(u) && u.WorksheetNOTo != null && u.WorksheetNOTo > maxWorksheetNOCompleted) maxWorksheetNOCompleted = u.WorksheetNOTo;
        });
        const resNotUpdate = res.NotUpdateMaxStudyScheduleIndex;
        const resNotWorksheet = res.NotUpdateMaxWorksheetNO;
        const notWorksheetFallback = maxWorksheetNODownloaded > 0 ? maxWorksheetNODownloaded : (maxWorksheetNOCompleted || null);
        return {
            SystemCountryCD: req.SystemCountryCD || 'USA', CenterID: req.CenterID || '', StudentID: req.StudentID || '', ClassID: req.ClassID || '',
            ClassStudentSeq: req.ClassStudentSeq != null ? req.ClassStudentSeq : null, SubjectCD: req.SubjectCD || '', WorksheetCD: req.WorksheetCD || '',
            DeleteSetInfoList: [], DiagnosticTestSetRegisterKbn: '0', FinishTestSetInfoList: [], InsertSetInfoList: [],
            NotDownloadLastUpdateTime: res.NotDownloadLastUpdateTime || null,
            NotUpdateMaxStudyScheduleIndex: (resNotUpdate != null ? resNotUpdate : maxIndexWithStudyDate),
            NotUpdateMaxWorksheetNO: (resNotWorksheet != null ? resNotWorksheet : notWorksheetFallback),
            client: req.client || CLIENT,
            id: lastApiResponseId != null ? String(lastApiResponseId + 1) : (req.id != null ? String(parseInt(req.id, 10) + 1) : String(Date.now()))
        };
    }
    function getEffectiveContext() {
        const fromStudy = buildPayloadFromStudyResult();
        if (fromStudy) return fromStudy;
        return lastRegisterStudySetRequest || null;
    }
    /** Get NotUpdateMaxStudyScheduleIndex and NotUpdateMaxWorksheetNO. Uses frozen study result for this Set when available. NotUpdateMaxWorksheetNO = highest worksheet already DOWNLOADED. */
    function getNotUpdateFromStudyResult(ctx) {
        let res = null;
        const key = ctx ? (ctx.StudentID || '') + '|' + (ctx.SubjectCD || '') + '|' + (ctx.WorksheetCD || '') : '';
        if (key.length > 2 && frozenStudyResultBySet[key]) res = getStudyResultData(frozenStudyResultBySet[key]);
        if (!res && lastStudyResult) res = getStudyResultData(lastStudyResult);
        if (!res) return null;
        const list = res.StudyUnitInfoList || [];
        let maxIndexWithStudyDate = 0, maxWorksheetNODownloaded = 0, maxWorksheetNOCompleted = 0;
        list.forEach(u => {
            if (u.StudyDate && u.StudyScheduleIndex != null && u.StudyScheduleIndex > maxIndexWithStudyDate) maxIndexWithStudyDate = u.StudyScheduleIndex;
            if (u.DownloadFlg === '1' && u.WorksheetNOTo != null && u.WorksheetNOTo > maxWorksheetNODownloaded) maxWorksheetNODownloaded = u.WorksheetNOTo;
            if (isStudyUnitCompleted(u) && u.WorksheetNOTo != null && u.WorksheetNOTo > maxWorksheetNOCompleted) maxWorksheetNOCompleted = u.WorksheetNOTo;
        });
        return {
            NotUpdateMaxStudyScheduleIndex: res.NotUpdateMaxStudyScheduleIndex != null ? res.NotUpdateMaxStudyScheduleIndex : maxIndexWithStudyDate,
            NotUpdateMaxWorksheetNO: res.NotUpdateMaxWorksheetNO != null ? res.NotUpdateMaxWorksheetNO : (maxWorksheetNODownloaded > 0 ? maxWorksheetNODownloaded : maxWorksheetNOCompleted)
        };
    }
    /** Next row only: we only add, never replace. Use max + 1 from StudyUnitInfoList. */
    function getNextStudyScheduleIndex(ctx, startPage, totalPages) {
        if (!ctx) return 1;
        const list = ctx.InsertSetInfoList;
        if (list && list.length) {
            let max = 0;
            list.forEach(x => { if (x.StudyScheduleIndex != null && x.StudyScheduleIndex > max) max = x.StudyScheduleIndex; });
            return max > 0 ? max + 1 : 1;
        }
        if (lastStudyResult) {
            const res = getStudyResultData(lastStudyResult);
            if (res && res.StudyUnitInfoList && res.StudyUnitInfoList.length) {
                let maxAll = 0;
                res.StudyUnitInfoList.forEach(u => { if (u.StudyScheduleIndex != null && u.StudyScheduleIndex > maxAll) maxAll = u.StudyScheduleIndex; });
                return maxAll > 0 ? maxAll + 1 : (res.NotUpdateMaxStudyScheduleIndex != null ? res.NotUpdateMaxStudyScheduleIndex : 0) + 1;
            }
        }
        const notUpdate = ctx.NotUpdateMaxStudyScheduleIndex != null ? ctx.NotUpdateMaxStudyScheduleIndex : 0;
        return notUpdate + 1;
    }
    function getWorksheetSetterContextLabel() {
        const ctx = getEffectiveContext();
        if (!ctx) return '(Open a student\u2019s Set page first)';
        const sid = ctx.StudentID || '';
        const subj = ctx.SubjectCD === '010' ? 'Math' : ctx.SubjectCD === '022' ? 'Reading' : (ctx.SubjectCD || '');
        const ws = ctx.WorksheetCD || '';
        let name = '';
        if (sid && capturedStudents.length) {
            const s = capturedStudents.find(st => (st.StudentID || st.LoginID) === sid);
            if (s) name = ' \u2013 ' + (s.FullName || s.StudentName || s.Name || '');
        }
        return sid + name + ' | ' + subj + ' ' + ws;
    }
    function updateWorksheetSetterUI() {
        const ctxEl = document.getElementById('kumon-bs-ctx');
        if (ctxEl) ctxEl.textContent = getWorksheetSetterContextLabel();
        const statusEl = document.getElementById('kumon-bs-status');
        if (statusEl) {
            const parts = [];
            parts.push(lastToken ? 'Token \u2713' : 'Token \u2717');
            parts.push(getEffectiveContext() ? 'Context \u2713' : 'Context \u2717');
            statusEl.textContent = parts.join(' ');
        }
        const hintEl = document.getElementById('kumon-bs-register-log-hint');
        if (hintEl) {
            if (registerLog.length === 0) hintEl.textContent = 'No RegisterStudySetInfo calls logged yet. Assign manually or use our Assign.';
            else {
                const last = registerLog[registerLog.length - 1];
                const res = last.response || {};
                hintEl.textContent = 'Last: ' + last.source + ' | ID: ' + (res.ID || res.id) + ' | ExclusionFlg: ' + (res.ExclusionFlg != null ? res.ExclusionFlg : '?') + ' | UpdateStudyInfoList: ' + ((res.UpdateStudyInfoList && res.UpdateStudyInfoList.length) || 0);
            }
        }
    }

    document.addEventListener('KumonBreakSetToken', function(ev) { const auth = ev.detail && ev.detail.authorization; if (auth) { lastToken = auth; updateWorksheetSetterUI(); } });
    document.addEventListener('KumonBreakSetRegister', function(ev) {
        const d = ev.detail;
        if (!d) return;
        try {
            if (d.requestBody) lastRegisterStudySetRequest = typeof d.requestBody === 'string' ? JSON.parse(d.requestBody) : d.requestBody;
            if (d.authorization) lastToken = d.authorization;
            updateWorksheetSetterUI();
        } catch (e) {}
    });
    document.addEventListener('KumonBreakSetStudyResult', function(ev) {
        const d = ev.detail;
        if (d && d.requestBody) lastStudyResultRequest = d.requestBody;
        if (d && d.responseData) lastStudyResult = d.responseData;
        if (d && d.responseData) {
            const req = (typeof d.requestBody === 'string' ? (() => { try { return JSON.parse(d.requestBody); } catch (e) { return null; } })() : d.requestBody) || {};
            const key = (req.StudentID || '') + '|' + (req.SubjectCD || '') + '|' + (req.WorksheetCD || '');
            if (key.length > 2 && !frozenStudyResultBySet[key]) {
                frozenStudyResultBySet[key] = d.responseData;
            }
        }
        if (d && d.responseData && typeof d.responseData === 'object') {
            const res = d.responseData.ID != null ? d.responseData : (d.responseData.data || d.responseData.Data || d.responseData);
            if (res && res.ID != null) { const n = parseInt(res.ID, 10); if (!isNaN(n)) lastApiResponseId = n; }
        }
        updateWorksheetSetterUI();
    });
    document.addEventListener('KumonStudyResultCapture', function(ev) {
        const d = ev.detail;
        if (!d) return;
        try {
            if (d.requestBody) lastStudyResultRequest = typeof d.requestBody === 'string' ? d.requestBody : JSON.stringify(d.requestBody);
            if (d.studyResultJson) lastStudyResult = JSON.parse(d.studyResultJson);
            updateWorksheetSetterUI();
        } catch (e) {}
    });
    document.addEventListener('KumonRegisterStudySetCapture', function(ev) {
        const d = ev.detail;
        if (!d) return;
        try {
            if (d.requestBody) lastRegisterStudySetRequest = typeof d.requestBody === 'string' ? JSON.parse(d.requestBody) : d.requestBody;
            if (d.authorization) lastToken = d.authorization;
            updateWorksheetSetterUI();
        } catch (e) {}
    });
    document.addEventListener('KumonTokenCapture', function(ev) {
        const auth = ev.detail && ev.detail.authorization;
        if (auth) { lastToken = auth; updateWorksheetSetterUI(); }
    });
    document.addEventListener('KumonStudyProfileCapture', function(ev) {
        try {
            const list = JSON.parse((ev.detail && ev.detail.studentsJson) || '[]');
            if (Array.isArray(list) && list.length) { capturedStudents = list; updateWorksheetSetterUI(); }
        } catch (e) {}
    });
    document.addEventListener('KumonBreakSetStudents', function(ev) {
        try {
            const list = JSON.parse((ev.detail && ev.detail.studentsJson) || '[]');
            if (Array.isArray(list) && list.length) { capturedStudents = list; updateWorksheetSetterUI(); }
        } catch (e) {}
    });
    document.addEventListener('KumonBreakSetRegisterLogUpdated', updateWorksheetSetterUI);

    const STUDY_PLAN_PRESETS_KEY = 'studyPlanPresets';
    function loadStudyPlanPresets(cb) {
        chrome.storage.local.get(STUDY_PLAN_PRESETS_KEY, function(o) {
            studyPlanPresets = Array.isArray(o[STUDY_PLAN_PRESETS_KEY]) ? o[STUDY_PLAN_PRESETS_KEY] : [];
            if (typeof cb === 'function') cb();
        });
    }
    function saveStudyPlanPresets() {
        chrome.storage.local.set({ [STUDY_PLAN_PRESETS_KEY]: studyPlanPresets }, function() {});
    }
    function renderStudyPlanPresets() {
        const listEl = document.getElementById('kumon-presets-list');
        if (!listEl) return;
        listEl.innerHTML = '';
        studyPlanPresets.forEach(function(p, i) {
            const endPage = (p.startPage || 1) + (p.totalPages || 1) - 1;
            const row = document.createElement('div');
            row.className = 'kumon-preset-row';
            row.innerHTML = '<span title="' + (p.name || '') + '">' + (p.name || 'Preset') + '</span> <span>p.' + (p.startPage || 1) + '\u2013' + endPage + ' ' + (p.patternKey || '') + '</span>';
            const applyBtn = document.createElement('button');
            applyBtn.type = 'button';
            applyBtn.textContent = 'Apply';
            applyBtn.dataset.index = String(i);
            applyBtn.addEventListener('click', function() { applyStudyPlanPreset(studyPlanPresets[parseInt(this.dataset.index, 10)]); });
            const delBtn = document.createElement('button');
            delBtn.type = 'button';
            delBtn.className = 'kumon-preset-del';
            delBtn.textContent = 'Del';
            delBtn.dataset.index = String(i);
            delBtn.addEventListener('click', function() {
                const idx = parseInt(this.dataset.index, 10);
                studyPlanPresets.splice(idx, 1);
                saveStudyPlanPresets();
                renderStudyPlanPresets();
            });
            row.appendChild(applyBtn);
            row.appendChild(delBtn);
            listEl.appendChild(row);
        });
    }
    function applyStudyPlanPreset(preset) {
        const statusEl = document.getElementById('kumon-presets-apply-status');
        function setStatus(msg, ok) {
            if (statusEl) { statusEl.textContent = msg || ''; statusEl.style.color = ok === true ? '#2e7d32' : (ok === false ? '#c62828' : ''); }
        }
        if (!preset || !preset.startPage || !preset.totalPages || !preset.patternKey || !PRESETS[preset.patternKey]) {
            setStatus('Invalid preset.', false);
            return;
        }
        const ctx = getEffectiveContext();
        if (!ctx) {
            setStatus('Open a student Set page first.', false);
            return;
        }
        if (!lastToken) {
            setStatus('No token. Navigate in the app first.', false);
            return;
        }
        const startPage = parseInt(preset.startPage, 10);
        const totalPages = parseInt(preset.totalPages, 10);
        if (isNaN(startPage) || isNaN(totalPages) || startPage < 1 || totalPages < 1) {
            setStatus('Invalid start/total.', false);
            return;
        }
        const patternSizes = PRESETS[preset.patternKey];
        const expanded = expandPatternToTotal(patternSizes, totalPages);
        if (!expanded) {
            setStatus('Pattern doesn\'t fit total pages.', false);
            return;
        }
        const truncatedTotal = expanded.reduce(function(a, b) { return a + b; }, 0);
        const nextIndex = getNextStudyScheduleIndex(ctx, startPage, truncatedTotal);
        const insertList = buildInsertSetInfoList(startPage, truncatedTotal, expanded, nextIndex);
        if (!insertList) {
            setStatus('Could not build sets.', false);
            return;
        }
        const payload = {};
        for (const k in ctx) if (Object.prototype.hasOwnProperty.call(ctx, k)) payload[k] = ctx[k];
        payload.InsertSetInfoList = insertList;
        payload.FinishTestSetInfoList = [];
        payload.DeleteSetInfoList = [];
        if (!payload.client || typeof payload.client !== 'object') payload.client = {};
        payload.client.applicationName = payload.client.applicationName || CLIENT.applicationName;
        payload.client.version = payload.client.version || CLIENT.version;
        payload.client.programName = payload.client.programName || CLIENT.programName;
        payload.client.os = payload.client.os || CLIENT.os;
        payload.client.machineName = payload.client.machineName != null ? payload.client.machineName : CLIENT.machineName;
        const notUpdate = getNotUpdateFromStudyResult(ctx);
        if (notUpdate) {
            payload.NotUpdateMaxStudyScheduleIndex = notUpdate.NotUpdateMaxStudyScheduleIndex;
            payload.NotUpdateMaxWorksheetNO = notUpdate.NotUpdateMaxWorksheetNO;
        }
        const nextId = lastApiResponseId != null ? lastApiResponseId + 1 : (ctx.id != null ? parseInt(ctx.id, 10) + 1 : null);
        payload.id = nextId != null ? String(nextId) : String(Date.now());
        setStatus('Sending...', null);
        fetch(REGISTER_STUDY_SET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': lastToken },
            body: JSON.stringify(payload)
        }).then(function(res) { return res.json(); }).then(function(data) {
            const rid = data && (data.ID != null ? data.ID : data.id);
            if (rid != null) { const n = parseInt(rid, 10); if (!isNaN(n)) lastApiResponseId = n; }
            registerLog.push({ ts: Date.now(), source: 'script', request: payload, response: data });
            if (registerLog.length > REGISTER_LOG_MAX) registerLog.shift();
            if (data && data.Result && data.Result.ResultCode === 0) {
                setStatus('Applied. ' + insertList.length + ' set(s).', true);
                log('Study plan preset applied: ' + (preset.name || '') + ' (' + insertList.length + ' set(s))', 'success');
            } else {
                const err = (data.Result && data.Result.Errors && data.Result.Errors[0]) || ('ResultCode=' + (data.Result && data.Result.ResultCode));
                setStatus('Error: ' + err, false);
            }
        }).catch(function(err) {
            setStatus('Request failed.', false);
        });
    }

    /** When the page sends RegisterStudySetInfo and a custom worksheet pattern is selected, build a replacement payload. Dispatch KumonPayloadReady so the page script can replace the request body. */
    document.addEventListener('KumonBuildReplacementPayload', function(ev) {
        let builtPayload = null;
        function done() { try { document.dispatchEvent(new CustomEvent('KumonPayloadReady', { detail: { payload: builtPayload } })); } catch (e) {} }
        if (!extensionEnabled) { done(); return; }
        const req = ev.detail && ev.detail.request;
        if (!req || typeof req !== 'object') { done(); return; }
        const patternKey = (ev.detail && ev.detail.patternKey) || '';
        if (!patternKey || !PRESETS[patternKey]) { done(); return; }
        const list = req.InsertSetInfoList;
        if (!list || !list.length) { done(); return; }
        let startPage = Infinity, endPage = -Infinity;
        list.forEach(function(item) {
            if (item.WorksheetNOFrom != null && item.WorksheetNOFrom < startPage) startPage = item.WorksheetNOFrom;
            if (item.WorksheetNOTo != null && item.WorksheetNOTo > endPage) endPage = item.WorksheetNOTo;
        });
        if (startPage === Infinity || endPage === -Infinity || endPage < startPage) { done(); return; }
        const totalPages = endPage - startPage + 1;
        const patternSizes = PRESETS[patternKey];
        const expanded = expandPatternToTotal(patternSizes, totalPages);
        if (!expanded) { done(); return; }
        const truncatedTotal = expanded.reduce((a, b) => a + b, 0);
        const pageStudyScheduleIndex = list[0].StudyScheduleIndex != null ? list[0].StudyScheduleIndex : getNextStudyScheduleIndex(req, startPage, truncatedTotal);
        const insertList = buildInsertSetInfoList(startPage, truncatedTotal, expanded, pageStudyScheduleIndex);
        if (!insertList) { done(); return; }
        const payload = {};
        for (const k in req) if (Object.prototype.hasOwnProperty.call(req, k)) payload[k] = req[k];
        payload.InsertSetInfoList = insertList;
        payload.FinishTestSetInfoList = req.FinishTestSetInfoList || [];
        payload.DeleteSetInfoList = req.DeleteSetInfoList || [];
        if (!payload.client || typeof payload.client !== 'object') payload.client = {};
        payload.client.applicationName = payload.client.applicationName || CLIENT.applicationName;
        payload.client.version = payload.client.version || CLIENT.version;
        payload.client.programName = payload.client.programName || CLIENT.programName;
        payload.client.os = payload.client.os || CLIENT.os;
        payload.client.machineName = payload.client.machineName != null ? payload.client.machineName : CLIENT.machineName;
        const nextId = lastApiResponseId != null ? lastApiResponseId + 1 : (req.id != null ? parseInt(req.id, 10) + 1 : null);
        payload.id = nextId != null ? String(nextId) : String(Date.now());
        builtPayload = payload;
        const actualEnd = startPage + truncatedTotal - 1;
        log('RegisterStudySetInfo replaced with pattern ' + patternKey + ' (' + insertList.length + ' set(s), p.' + startPage + '\u2013' + actualEnd + (actualEnd < endPage ? ', truncated from ' + endPage : '') + ')', 'success');
        done();
    });

    // Create UI
    function createUI() {
        // Remove existing UI if present (unified or legacy panel)
        const existing = document.getElementById('kumon-extensions-ui') || document.getElementById('kumon-auto-grader-ui');
        if (existing) existing.remove();

        // Create container (unified panel with functionality dropdown)
        uiContainer = document.createElement('div');
        uiContainer.id = 'kumon-extensions-ui';
        uiContainer.innerHTML = `
            <div class="kumon-grader-header">
                <select id="kumon-tool-select" class="kumon-tool-select" title="Switch tool">
                    <option value="grader">Auto Grader</option>
                    <option value="presets">Study plan presets</option>
                </select>
                <div class="header-buttons">
                    <button class="collapse-panel" id="kumon-collapse" title="Collapse">▾</button>
                    <button class="toggle-settings" id="toggle-settings" title="Settings">⚙️</button>
                    <button class="resize-handle" id="resize-handle" title="Resize">⛶</button>
                </div>
            </div>
            <div id="kumon-panel-grader" class="kumon-panel">
            <div class="kumon-grader-body">
                <div class="status-section">
                    <div class="status-text status-info" id="status-text">Ready</div>
                    <div class="progress-container">
                        <div class="progress-bar" id="progress-bar"></div>
                    </div>
                    <div class="progress-text" id="progress-text">0 / 0</div>
                </div>
                <div class="mode-section">
                    <label class="mode-label">Mode:</label>
                    <div class="mode-buttons">
                        <button class="mode-btn ${CONFIG.mode === 'x' ? 'active' : ''}" data-mode="x" title="Mark as X">✗</button>
                        <button class="mode-btn ${CONFIG.mode === 'triangle' ? 'active' : ''}" data-mode="triangle" title="Mark as Triangle">△</button>
                        <button class="mode-btn ${CONFIG.mode === 'clear' ? 'active' : ''}" data-mode="clear" title="Clear marks">○</button>
                        <button class="mode-btn ${CONFIG.mode === 'reapply' ? 'active' : ''}" data-mode="reapply" title="Reapply from results">↻</button>
                    </div>
                </div>
                <div class="controls-section">
                    <button class="btn btn-primary" id="start-btn">▶️ Start</button>
                    <button class="btn btn-danger" id="stop-btn" disabled>⏹️ Stop</button>
                    <button class="btn btn-secondary" id="refresh-btn">🔄 Refresh</button>
                </div>
                <div class="info-section">
                    <div class="info-item">
                        <span class="info-label">Hotkey:</span>
                        <span class="info-value">Alt + R</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Boxes Found:</span>
                        <span class="info-value" id="boxes-count">0</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Result Marks:</span>
                        <span class="info-value" id="result-count">0</span>
                    </div>
                </div>
            </div>
            <div class="kumon-grader-settings" id="settings-panel" style="display: none;">
                <div class="settings-header">Settings</div>
                <div class="settings-content">
                    <div class="setting-item">
                        <label>
                            <input type="checkbox" id="restrict-visible" ${CONFIG.restrictToVisible ? 'checked' : ''}>
                            Restrict to visible worksheet
                        </label>
                    </div>
                    <div class="setting-item">
                        <label>
                            Click Delay (ms):
                            <input type="number" id="click-delay" value="${CONFIG.clickDelay}" min="0" max="500" step="10">
                        </label>
                    </div>
                    <div class="setting-item">
                        <label>
                            Clicks Per Box:
                            <select id="clicks-per-box">
                                <option value="1" ${CONFIG.clicksPerBox === 1 ? 'selected' : ''}>1 (X)</option>
                                <option value="2" ${CONFIG.clicksPerBox === 2 ? 'selected' : ''}>2 (Triangle)</option>
                                <option value="3" ${CONFIG.clicksPerBox === 3 ? 'selected' : ''}>3 (Clear)</option>
                            </select>
                        </label>
                    </div>
                    <div class="setting-item">
                        <label>
                            <input type="checkbox" id="enable-logging" ${CONFIG.enableLogging ? 'checked' : ''}>
                            Enable console logging
                        </label>
                    </div>
                </div>
            </div>
            </div>
            <div id="kumon-panel-presets" class="kumon-panel" style="display: none;">
                <div class="kumon-presets-body">
                    <div class="kumon-presets-label">Study plan presets</div>
                    <div id="kumon-presets-list" class="kumon-presets-list"></div>
                    <div class="kumon-presets-add">
                        <input type="text" id="kumon-preset-name" placeholder="Name" class="kumon-preset-input" />
                        <input type="number" id="kumon-preset-start" min="1" placeholder="Start" class="kumon-preset-input kumon-preset-num" />
                        <input type="number" id="kumon-preset-total" min="1" placeholder="Total" class="kumon-preset-input kumon-preset-num" />
                        <select id="kumon-preset-pattern" class="kumon-preset-select">
                            <option value="10">10</option>
                            <option value="5">5</option>
                            <option value="4-3-3">4-3-3</option>
                            <option value="3-2">3-2</option>
                            <option value="2-2">2-2</option>
                        </select>
                        <button type="button" id="kumon-preset-add-btn" class="kumon-preset-btn">Add</button>
                    </div>
                    <div id="kumon-presets-apply-status" class="kumon-presets-status"></div>
                </div>
            </div>
        `;

        // Add styles
        const style = document.createElement('style');
        style.textContent = `
            #kumon-extensions-ui {
                position: fixed;
                top: 20px;
                right: 20px;
                width: 280px;
                min-width: 240px;
                max-width: 420px;
                min-height: 200px;
                background: linear-gradient(180deg, #02497e 0%, #013257 100%);
                border-radius: 8px;
                box-shadow: 0 8px 32px rgba(1, 50, 87, 0.4);
                z-index: 10000;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                color: #333;
                overflow: hidden;
                transition: transform 0.3s ease;
                resize: both;
                user-select: none;
                font-size: 12px;
                border: 1px solid #d0dde2;
            }

            #kumon-extensions-ui:hover {
                transform: translateY(-2px);
                box-shadow: 0 12px 40px rgba(1, 50, 87, 0.5);
            }

            #kumon-extensions-ui.kumon-collapsed {
                min-height: 0;
                height: auto;
            }

            #kumon-extensions-ui.kumon-collapsed .kumon-panel,
            #kumon-extensions-ui.kumon-collapsed .kumon-grader-settings {
                display: none !important;
            }

            #kumon-extensions-ui.kumon-collapsed .resize-handle {
                display: none;
            }

            .kumon-grader-header {
                padding: 8px 12px;
                background: #013257;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid #d0dde2;
                cursor: move;
                color: white;
            }

            .kumon-grader-header h3 {
                margin: 0;
                font-size: 13px;
                font-weight: 600;
                flex: 0 0 auto;
                color: white;
            }

            .header-buttons {
                display: flex;
                gap: 6px;
                align-items: center;
            }

            .kumon-tool-select {
                padding: 4px 6px;
                font-size: 11px;
                border-radius: 4px;
                border: 1px solid #d0dde2;
                background: #fff;
                color: #333;
                cursor: pointer;
            }
            .kumon-tool-select option { background: #fff; color: #333; }

            .kumon-panel { overflow: auto; background: #f8fafb; color: #333; }

            .toggle-settings, .resize-handle, .collapse-panel {
                background: #02497e;
                border: 1px solid #d0dde2;
                color: white;
                padding: 4px 6px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 11px;
                transition: all 0.2s;
                line-height: 1;
            }

            .toggle-settings:hover, .resize-handle:hover, .collapse-panel:hover {
                background: #013257;
                transform: scale(1.05);
            }

            .resize-handle {
                cursor: nwse-resize;
            }

            .kumon-grader-body {
                padding: 12px;
                background: #f8fafb;
                color: #333;
            }

            .status-section {
                margin-bottom: 10px;
                display: block;
                width: 100%;
            }

            .status-text {
                font-size: 11px;
                font-weight: 500;
                margin-bottom: 6px;
                padding: 6px 8px;
                border-radius: 4px;
                background: #fff;
                border: 1px solid #d0dde2;
                color: #333;
                text-align: center;
                display: block;
                width: 100%;
            }

            .status-info {
                background: #e6f4ff;
                border-color: #02497e;
            }

            .status-processing {
                background: #fff3e0;
                border-color: #f57c00;
                animation: pulse 1.5s ease-in-out infinite;
            }

            .status-success {
                background: #e8f5e9;
                border-color: #2e7d32;
            }

            .status-warning {
                background: #fff3e0;
                border-color: #f57c00;
            }

            .status-stopped {
                background: #f5f5f5;
                border-color: #9e9e9e;
            }

            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.7; }
            }

            .progress-container {
                width: 100%;
                height: 8px;
                background: #e0e8ec;
                border-radius: 4px;
                overflow: hidden;
                margin-bottom: 8px;
            }

            .progress-bar {
                height: 100%;
                background: linear-gradient(90deg, #02497e, #013257);
                width: 0%;
                transition: width 0.3s ease;
                border-radius: 4px;
            }

            .progress-text {
                text-align: center;
                font-size: 10px;
                color: #333;
                display: block;
                width: 100%;
                margin-bottom: 8px;
            }

            .mode-section {
                margin-bottom: 10px;
                display: block;
                width: 100%;
                clear: both;
            }

            .mode-label {
                display: block;
                font-size: 10px;
                color: #555;
                margin-bottom: 6px;
                font-weight: 500;
            }

            .mode-buttons {
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 4px;
            }

            .mode-btn {
                padding: 6px 4px;
                border: 2px solid #d0dde2;
                border-radius: 4px;
                font-size: 10px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s;
                color: #333;
                background: #fff;
            }

            .mode-btn:hover {
                background: #e6f4ff;
                border-color: #02497e;
                transform: translateY(-1px);
            }

            .mode-btn.active {
                background: #02497e;
                border-color: #013257;
                color: white;
                box-shadow: 0 2px 8px rgba(1, 50, 87, 0.3);
            }

            .controls-section {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 6px;
                margin-bottom: 10px;
            }

            .btn {
                padding: 8px 10px;
                border: 1px solid #d0dde2;
                border-radius: 5px;
                font-size: 11px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s;
                color: #333;
                background: #fff;
                box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            }

            .btn:disabled {
                opacity: 0.5;
                cursor: not-allowed;
                transform: none !important;
                box-shadow: none !important;
            }

            .btn-primary {
                background: #02497e;
                color: white;
                border-color: #013257;
            }

            .btn-primary:hover:not(:disabled) {
                background: #013257;
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(1, 50, 87, 0.35);
            }

            .btn-primary:active:not(:disabled) {
                transform: translateY(0);
            }

            .btn-danger {
                background: #c62828;
                color: white;
                border-color: #b71c1c;
            }

            .btn-danger:hover:not(:disabled) {
                background: #b71c1c;
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(198, 40, 40, 0.35);
            }

            .btn-danger:active:not(:disabled) {
                transform: translateY(0);
            }

            .btn-secondary {
                background: #fff;
                color: #02497e;
                border-color: #02497e;
            }

            .btn-secondary:hover:not(:disabled) {
                background: #e6f4ff;
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(2, 73, 126, 0.2);
            }

            .btn-secondary:active:not(:disabled) {
                transform: translateY(0);
            }

            .info-section {
                background: #fff;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                padding: 8px;
                color: #333;
            }

            .info-item {
                display: flex;
                justify-content: space-between;
                margin-bottom: 4px;
                font-size: 10px;
            }

            .info-item:last-child {
                margin-bottom: 0;
            }

            .info-label {
                color: #555;
            }

            .info-value {
                font-weight: 600;
                color: #333;
            }

            .kumon-presets-body {
                padding: 12px;
                background: #f8fafb;
                color: #333;
            }
            .kumon-presets-label {
                font-size: 11px;
                font-weight: 600;
                margin-bottom: 6px;
                color: #333;
            }
            .kumon-presets-list {
                max-height: 120px;
                overflow-y: auto;
                margin-bottom: 8px;
            }
            .kumon-preset-row {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 6px;
                padding: 4px 6px;
                background: #fff;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                margin-bottom: 4px;
                font-size: 10px;
                color: #333;
            }
            .kumon-preset-row span {
                flex: 1;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
            }
            .kumon-preset-row button {
                flex-shrink: 0;
                padding: 2px 6px;
                font-size: 10px;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                cursor: pointer;
                background: #fff;
                color: #02497e;
            }
            .kumon-preset-row button:hover {
                background: #e6f4ff;
            }
            .kumon-preset-row button.kumon-preset-del {
                background: #ffebee;
                color: #c62828;
                border-color: #c62828;
            }
            .kumon-presets-add {
                display: flex;
                flex-wrap: wrap;
                gap: 4px;
                align-items: center;
            }
            .kumon-preset-input, .kumon-preset-select {
                padding: 4px 6px;
                font-size: 10px;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                background: #fff;
                color: #333;
            }
            .kumon-preset-input { width: 60px; }
            .kumon-preset-input.kumon-preset-num { width: 44px; }
            .kumon-preset-select { width: 56px; }
            #kumon-preset-name { width: 72px; }
            .kumon-preset-btn {
                padding: 4px 8px;
                font-size: 10px;
                border: 1px solid #02497e;
                border-radius: 4px;
                cursor: pointer;
                background: #02497e;
                color: white;
            }
            .kumon-preset-btn:hover {
                background: #013257;
            }
            .kumon-presets-status {
                margin-top: 4px;
                font-size: 10px;
                min-height: 14px;
                color: #555;
            }

            .kumon-grader-settings {
                border-top: 1px solid #d0dde2;
                padding: 10px 12px;
                background: #f0f4f6;
                color: #333;
            }

            .settings-header {
                font-weight: 600;
                margin-bottom: 8px;
                font-size: 11px;
                color: #333;
            }

            .settings-content {
                display: flex;
                flex-direction: column;
                gap: 8px;
            }

            .setting-item {
                font-size: 11px;
                color: #333;
            }

            .setting-item label {
                display: flex;
                align-items: center;
                gap: 8px;
                cursor: pointer;
            }

            .setting-item input[type="checkbox"] {
                width: 18px;
                height: 18px;
                cursor: pointer;
            }

            .setting-item input[type="number"] {
                width: 80px;
                padding: 4px 8px;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                margin-left: 8px;
                font-size: 12px;
                background: #fff;
                color: #333;
            }

            .setting-item select {
                width: 100px;
                padding: 4px 8px;
                border: 1px solid #d0dde2;
                border-radius: 4px;
                margin-left: 8px;
                font-size: 12px;
                background: #fff;
                color: #333;
            }
            .kumon-bs-body { padding: 12px; background: #f8fafb; color: #333; }
            .kumon-bs-block { margin-bottom: 12px; }
            .kumon-bs-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #013257; margin-bottom: 4px; }
            .kumon-bs-ctx { font-size: 12px; color: #333; line-height: 1.4; margin-bottom: 4px; }
            .kumon-bs-hint { font-size: 10px; color: #555; }
            .kumon-bs-input { width: 100%; padding: 8px 10px; font-size: 13px; background: #fff; color: #333; border: 1px solid #d0dde2; border-radius: 8px; margin-bottom: 4px; box-sizing: border-box; }
            .kumon-bs-presets { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
            .kumon-bs-preset { padding: 6px 10px; font-size: 11px; border-radius: 6px; border: 1px solid #02497e; background: #fff; color: #02497e; cursor: pointer; }
            .kumon-bs-preset:hover { background: #e6f4ff; }
            .kumon-bs-preview { font-size: 11px; font-family: monospace; background: #fff; padding: 8px; border-radius: 6px; margin: 8px 0; color: #333; border: 1px solid #d0dde2; white-space: pre-wrap; word-break: break-all; min-height: 1.2em; }
            .kumon-bs-btn { width: 100%; padding: 10px 14px; background: #02497e; color: #fff; border: 1px solid #013257; border-radius: 10px; font-weight: 600; font-size: 13px; cursor: pointer; }
            .kumon-bs-btn:hover { background: #013257; }
            .kumon-bs-result { font-size: 12px; margin-top: 8px; min-height: 1.4em; color: #333; }
            .kumon-bs-details { margin-top: 10px; border: 1px solid #d0dde2; border-radius: 8px; overflow: hidden; background: #fff; }
            .kumon-bs-details summary { padding: 6px 10px; font-size: 11px; font-weight: 600; cursor: pointer; background: #f0f4f6; color: #013257; }
            .kumon-bs-textarea { width: 100%; min-height: 80px; margin: 6px 0; padding: 8px; font-size: 10px; font-family: monospace; background: #fff; color: #333; border: 1px solid #d0dde2; border-radius: 6px; resize: vertical; box-sizing: border-box; }
            .kumon-bs-btn-secondary { width: 100%; padding: 8px; font-size: 12px; border-radius: 8px; border: 1px solid #02497e; background: #fff; color: #02497e; cursor: pointer; }
        `;

        document.head.appendChild(style);
        document.body.appendChild(uiContainer);

        // Get UI elements
        statusText = document.getElementById('status-text');
        progressBar = document.getElementById('progress-bar');
        progressText = document.getElementById('progress-text');
        startButton = document.getElementById('start-btn');
        stopButton = document.getElementById('stop-btn');
        const refreshButton = document.getElementById('refresh-btn');
        const toggleSettings = document.getElementById('toggle-settings');
        const resizeHandle = document.getElementById('resize-handle');
        const collapseButton = document.getElementById('kumon-collapse');
        settingsPanel = document.getElementById('settings-panel');
        const restrictVisible = document.getElementById('restrict-visible');
        const clickDelayInput = document.getElementById('click-delay');
        const clicksPerBoxSelect = document.getElementById('clicks-per-box');
        const enableLogging = document.getElementById('enable-logging');
        const boxesCount = document.getElementById('boxes-count');
        const resultCount = document.getElementById('result-count');
        const modeButtons = document.querySelectorAll('.mode-btn');

        // Make UI draggable
        let isDragging = false;
        let dragOffset = { x: 0, y: 0 };
        const header = uiContainer.querySelector('.kumon-grader-header');

        header.addEventListener('mousedown', (e) => {
            if (e.target.classList.contains('toggle-settings') ||
                e.target.classList.contains('resize-handle') ||
                e.target.id === 'kumon-tool-select' ||
                (e.target.closest && (e.target.closest('.header-buttons') || e.target.closest('#kumon-tool-select')))) {
                return;
            }
            isDragging = true;
            const rect = uiContainer.getBoundingClientRect();
            dragOffset.x = e.clientX - rect.left;
            dragOffset.y = e.clientY - rect.top;
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (isDragging) {
                uiContainer.style.left = (e.clientX - dragOffset.x) + 'px';
                uiContainer.style.top = (e.clientY - dragOffset.y) + 'px';
                uiContainer.style.right = 'auto';
            }
        });

        document.addEventListener('mouseup', () => {
            isDragging = false;
        });

        // Make UI resizable
        let isResizing = false;
        let resizeStart = { x: 0, y: 0, width: 0, height: 0 };

        resizeHandle.addEventListener('mousedown', (e) => {
            isResizing = true;
            const rect = uiContainer.getBoundingClientRect();
            resizeStart.x = e.clientX;
            resizeStart.y = e.clientY;
            resizeStart.width = rect.width;
            resizeStart.height = rect.height;
            e.preventDefault();
            e.stopPropagation();
        });

        document.addEventListener('mousemove', (e) => {
            if (isResizing) {
                const deltaX = e.clientX - resizeStart.x;
                const deltaY = e.clientY - resizeStart.y;
                const newWidth = Math.max(320, Math.min(600, resizeStart.width + deltaX));
                const newHeight = Math.max(300, resizeStart.height + deltaY);
                uiContainer.style.width = newWidth + 'px';
                uiContainer.style.height = newHeight + 'px';
            }
        });

        document.addEventListener('mouseup', () => {
            isResizing = false;
        });

        // Mode button listeners
        modeButtons.forEach(btn => {
            btn.addEventListener('click', () => {
                modeButtons.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                CONFIG.mode = btn.dataset.mode;
                log(`Mode changed to: ${CONFIG.mode}`, 'info');
                refreshButton.click();
            });
        });

        // Event listeners
        startButton.addEventListener('click', () => { if (!extensionEnabled) return; processAllBoxes(); });
        stopButton.addEventListener('click', stopProcessing);
        refreshButton.addEventListener('click', () => {
            const boxes = findFailureBoxes();
            const resultList = findResultBoxes();
            boxesCount.textContent = boxes.length;
            resultCount.textContent = resultList ? resultList.length : 0;

            if (CONFIG.mode === 'reapply') {
                const count = resultList ? resultList.length : 0;
                updateStatus(`Found ${count} result marks to reapply`, 'info');
                log(`Refreshed: Found ${count} result marks`, 'info');
            } else {
                updateStatus(`Found ${boxes.length} failure boxes`, 'info');
                log(`Refreshed: Found ${boxes.length} boxes`, 'info');
            }
        });

        toggleSettings.addEventListener('click', () => {
            const isVisible = settingsPanel.style.display !== 'none';
            settingsPanel.style.display = isVisible ? 'none' : 'block';
        });

        if (collapseButton) {
            let isCollapsed = false;
            collapseButton.addEventListener('click', (e) => {
                e.stopPropagation();
                isCollapsed = !isCollapsed;
                uiContainer.classList.toggle('kumon-collapsed', isCollapsed);
                collapseButton.textContent = isCollapsed ? '▸' : '▾';
                collapseButton.title = isCollapsed ? 'Expand' : 'Collapse';
            });
        }

        restrictVisible.addEventListener('change', (e) => {
            CONFIG.restrictToVisible = e.target.checked;
        });

        clickDelayInput.addEventListener('change', (e) => {
            CONFIG.clickDelay = parseInt(e.target.value) || 50;
        });

        clicksPerBoxSelect.addEventListener('change', (e) => {
            CONFIG.clicksPerBox = parseInt(e.target.value) || 1;
        });

        enableLogging.addEventListener('change', (e) => {
            CONFIG.enableLogging = e.target.checked;
        });

        loadStudyPlanPresets(renderStudyPlanPresets);
        const toolSelect = document.getElementById('kumon-tool-select');
        const panelGrader = document.getElementById('kumon-panel-grader');
        const panelPresets = document.getElementById('kumon-panel-presets');
        if (toolSelect && panelGrader && panelPresets) {
            toolSelect.addEventListener('change', function() {
                const v = toolSelect.value;
                panelGrader.style.display = v === 'grader' ? 'block' : 'none';
                panelPresets.style.display = v === 'presets' ? 'block' : 'none';
            });
        }
        document.getElementById('kumon-preset-add-btn').addEventListener('click', function() {
            const nameEl = document.getElementById('kumon-preset-name');
            const startEl = document.getElementById('kumon-preset-start');
            const totalEl = document.getElementById('kumon-preset-total');
            const patternEl = document.getElementById('kumon-preset-pattern');
            const name = (nameEl && nameEl.value || '').trim() || 'Preset';
            const start = parseInt(startEl && startEl.value, 10);
            const total = parseInt(totalEl && totalEl.value, 10);
            const patternKey = patternEl && patternEl.value || '4-3-3';
            if (isNaN(start) || isNaN(total) || start < 1 || total < 1 || !PRESETS[patternKey]) return;
            studyPlanPresets.push({
                id: Date.now().toString(36) + Math.random().toString(36).slice(2),
                name: name,
                startPage: start,
                totalPages: total,
                patternKey: patternKey
            });
            saveStudyPlanPresets();
            renderStudyPlanPresets();
            if (nameEl) nameEl.value = '';
            if (startEl) startEl.value = '';
            if (totalEl) totalEl.value = '';
        });

        // Initial refresh
        refreshButton.click();
    }

    function tearDownWorksheetSetter() {
        document.querySelectorAll('[data-kumon-pattern]').forEach(el => el.remove());
        const styleEl = document.getElementById('kumon-ws-pattern-style');
        if (styleEl) styleEl.remove();
        window.__kumonWorksheetPerStudyHooked = false;
        if (document.body) {
            document.body.removeAttribute('data-kumon-ws-pattern');
            document.body.dataset.kumonWorksheetPattern = '';
        }
        window.__kumonWorksheetPattern = undefined;
    }

    function ensureToggleButton() {
        let wrap = document.getElementById('kumon-extension-toggle-wrap');
        if (wrap) return;
        wrap = document.createElement('div');
        wrap.id = 'kumon-extension-toggle-wrap';
        wrap.style.cssText = 'position:fixed;bottom:16px;right:16px;z-index:10001;';
        const btn = document.createElement('button');
        btn.id = 'kumon-extension-toggle-btn';
        btn.type = 'button';
        btn.style.cssText = 'padding:6px 12px;font-size:12px;border-radius:6px;border:1px solid #013257;background:#02497e;color:#fff;cursor:pointer;font-family:inherit;box-shadow:0 2px 8px rgba(1,50,87,0.25);';
        function updateBtnText() { btn.textContent = extensionEnabled ? 'Disable Extension' : 'Enable Extension'; }
        updateBtnText();
        btn.addEventListener('click', () => {
            extensionEnabled = !extensionEnabled;
            chrome.storage.local.set({ extensionEnabled }, function() {});
            syncBodyDataset();
            const panel = document.getElementById('kumon-extensions-ui');
            if (panel) panel.style.display = extensionEnabled ? '' : 'none';
            if (extensionEnabled) {
                injectWorksheetPerStudyOptions();
            } else {
                tearDownWorksheetSetter();
            }
            updateBtnText();
        });
        wrap.appendChild(btn);
        document.body.appendChild(wrap);
    }

    function updateToggleButtonText() {
        const btn = document.getElementById('kumon-extension-toggle-btn');
        if (btn) btn.textContent = extensionEnabled ? 'Disable Extension' : 'Enable Extension';
    }

    function injectWorksheetPerStudyOptions() {
        if (window.__kumonWorksheetPerStudyHooked) return;
        window.__kumonWorksheetPerStudyHooked = true;

        const style = document.createElement('style');
        style.id = 'kumon-ws-pattern-style';
        style.textContent = [
            'body[data-kumon-ws-pattern] .ATD0010P-root .menu-bar .menu-right .setting-container .options.setting-options .option.option-select,',
            'body[data-kumon-ws-pattern] .setting-container .options.setting-options .option.option-select { background: #fff !important; }',
            'body[data-kumon-ws-pattern="4-3-3"] .ATD0010P-root .menu-bar .menu-right .setting-container .options.setting-options .option[data-kumon-pattern="4-3-3"],',
            'body[data-kumon-ws-pattern="4-3-3"] .setting-container .options.setting-options .option[data-kumon-pattern="4-3-3"] { background: #e6f4ff !important; }',
            'body[data-kumon-ws-pattern="3-2"] .ATD0010P-root .menu-bar .menu-right .setting-container .options.setting-options .option[data-kumon-pattern="3-2"],',
            'body[data-kumon-ws-pattern="3-2"] .setting-container .options.setting-options .option[data-kumon-pattern="3-2"] { background: #e6f4ff !important; }',
            'body[data-kumon-ws-pattern="2-2"] .ATD0010P-root .menu-bar .menu-right .setting-container .options.setting-options .option[data-kumon-pattern="2-2"],',
            'body[data-kumon-ws-pattern="2-2"] .setting-container .options.setting-options .option[data-kumon-pattern="2-2"] { background: #e6f4ff !important; }'
        ].join(' ');
        (document.head || document.documentElement).appendChild(style);

        document.addEventListener('click', function clearCustomSelectionOnNativeOption(e) {
            const opt = e.target.closest('.setting-container .options.setting-options .option');
            if (!opt) return;
            if (opt.dataset.kumonPattern) return;
            document.body.removeAttribute('data-kumon-ws-pattern');
            if (document.body) document.body.dataset.kumonWorksheetPattern = '';
        }, true);

        const LABELS = [
            '4-3-3 worksheets per study',
            '3-2 worksheets per study',
            '2-2 worksheets per study'
        ];
        const KEYS = ['4-3-3', '3-2', '2-2'];

        const extendAll = () => {
            if (!extensionEnabled) return;
            const optionContainers = document.querySelectorAll('.setting-container .options.setting-options');
            if (!optionContainers.length) return;

            optionContainers.forEach(optionsEl => {
                // Collect existing labels from both .option elements and raw text nodes
                const elementLabels = Array.from(optionsEl.querySelectorAll('.option.setting-options')).map(el => (el.textContent || '').trim());
                const textLabels = Array.from(optionsEl.childNodes)
                    .filter(n => n.nodeType === Node.TEXT_NODE)
                    .map(n => (n.textContent || '').trim())
                    .filter(Boolean);
                const existing = new Set([...elementLabels, ...textLabels]);

                const template = optionsEl.querySelector('.option.setting-options') || optionsEl.firstElementChild;
                let baseClass = 'option setting-options';
                if (template && template.className) {
                    // Strip any "option-select" class so new options are not always highlighted
                    baseClass = template.className.replace(/\boption-select\b/g, '').trim() || 'option setting-options';
                }

                LABELS.forEach((label, idx) => {
                    const key = KEYS[idx] || label;
                    if (existing.has(label)) return;
                    const opt = document.createElement('div');
                    opt.className = baseClass;
                    opt.textContent = label;
                    opt.dataset.kumonPattern = key;
                    opt.addEventListener('click', (e) => {
                        // Remember our own selected pattern and just close the menu.
                        // Let the site handle its own visual selection (option-select).
                        e.stopPropagation();
                        window.__kumonWorksheetPattern = key;
                        if (document.body) document.body.dataset.kumonWorksheetPattern = key;
                        const container = opt.closest('.options.setting-options');
                        if (container) {
                            container.setAttribute('hidden', '');
                        }
                    });
                    optionsEl.appendChild(opt);
                });
            });
        };

        // Try immediately (in case menu is already rendered), and on every click
        extendAll();
        document.addEventListener('click', extendAll, true);
    }

    // Hotkey handler
    function handleHotkey(event) {
        if (!extensionEnabled) return;
        if (event.altKey && event.key === 'r' && !event.ctrlKey && !event.shiftKey) {
            event.preventDefault();
            if (!isProcessing) {
                processAllBoxes();
            } else {
                stopProcessing();
            }
        }
    }

    // Initialize
    function init() {
        injectPageScript();
        log('Kumon Extensions initialized', 'success');
        syncBodyDataset();
        chrome.storage.local.get({ extensionEnabled: true }, function(items) {
            extensionEnabled = items.extensionEnabled !== false;
            createUI();
            ensureToggleButton();
            if (extensionEnabled) {
                injectWorksheetPerStudyOptions();
            } else {
                tearDownWorksheetSetter();
                if (uiContainer) uiContainer.style.display = 'none';
            }
            updateToggleButtonText();
            syncBodyDataset();
        });
        document.addEventListener('keydown', handleHotkey);
    }

    // Wait for DOM to be ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    // Re-initialize on navigation (for SPAs)
    let lastUrl = location.href;
    new MutationObserver(() => {
        const url = location.href;
        if (url !== lastUrl) {
            lastUrl = url;
            setTimeout(init, 1000);
        }
    }).observe(document, { subtree: true, childList: true });

})();

