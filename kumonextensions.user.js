// ==UserScript==
// @name         kumonextensions
// @namespace    https://github.com/Invisibl5/kumonextensions
// @version      0.2.1
// @description  Kumon Auto Grader (X / Triangle / Clear + Reapply)
// @author       Invisibl5
// @match        *://*/*
// @run-at       document-idle
// @grant        none
// @updateURL    https://raw.githubusercontent.com/Invisibl5/kumonextensions/main/kumonextensions.user.js
// @downloadURL  https://raw.githubusercontent.com/Invisibl5/kumonextensions/main/kumonextensions.user.js
// ==/UserScript==

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

    // Create UI
    function createUI() {
        // Remove existing UI if present
        const existing = document.getElementById('kumon-auto-grader-ui');
        if (existing) {
            existing.remove();
        }

        // Create container
        uiContainer = document.createElement('div');
        uiContainer.id = 'kumon-auto-grader-ui';
        uiContainer.innerHTML = `
            <div class="kumon-grader-header">
                <h3>🎯 Kumon Auto Grader</h3>
                <div class="header-buttons">
                    <button class="toggle-settings" id="toggle-settings" title="Settings">⚙️</button>
                    <button class="resize-handle" id="resize-handle" title="Resize">⛶</button>
                </div>
            </div>
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
        `;

        // Add styles
        const style = document.createElement('style');
        style.textContent = `
            #kumon-auto-grader-ui {
                position: fixed;
                top: 20px;
                right: 20px;
                width: 240px;
                min-width: 200px;
                max-width: 400px;
                min-height: 200px;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                border-radius: 8px;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
                z-index: 10000;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                color: white;
                overflow: hidden;
                transition: transform 0.3s ease;
                resize: both;
                user-select: none;
                font-size: 12px;
            }

            #kumon-auto-grader-ui:hover {
                transform: translateY(-2px);
                box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4);
            }

            .kumon-grader-header {
                padding: 8px 12px;
                background: rgba(255, 255, 255, 0.1);
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid rgba(255, 255, 255, 0.2);
                cursor: move;
            }

            .kumon-grader-header h3 {
                margin: 0;
                font-size: 13px;
                font-weight: 600;
                flex: 1;
            }

            .header-buttons {
                display: flex;
                gap: 6px;
            }

            .toggle-settings, .resize-handle {
                background: rgba(255, 255, 255, 0.2);
                border: none;
                color: white;
                padding: 4px 6px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 11px;
                transition: all 0.2s;
                line-height: 1;
            }

            .toggle-settings:hover, .resize-handle:hover {
                background: rgba(255, 255, 255, 0.3);
                transform: scale(1.05);
            }

            .resize-handle {
                cursor: nwse-resize;
            }

            .kumon-grader-body {
                padding: 12px;
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
                background: rgba(255, 255, 255, 0.1);
                text-align: center;
                display: block;
                width: 100%;
            }

            .status-info {
                background: rgba(33, 150, 243, 0.3);
            }

            .status-processing {
                background: rgba(255, 152, 0, 0.3);
                animation: pulse 1.5s ease-in-out infinite;
            }

            .status-success {
                background: rgba(76, 175, 80, 0.3);
            }

            .status-warning {
                background: rgba(255, 152, 0, 0.3);
            }

            .status-stopped {
                background: rgba(158, 158, 158, 0.3);
            }

            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.7; }
            }

            .progress-container {
                width: 100%;
                height: 8px;
                background: rgba(255, 255, 255, 0.2);
                border-radius: 4px;
                overflow: hidden;
                margin-bottom: 8px;
            }

            .progress-bar {
                height: 100%;
                background: linear-gradient(90deg, #4CAF50, #8BC34A);
                width: 0%;
                transition: width 0.3s ease;
                border-radius: 4px;
            }

            .progress-text {
                text-align: center;
                font-size: 10px;
                opacity: 0.9;
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
                opacity: 0.9;
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
                border: 2px solid rgba(255, 255, 255, 0.3);
                border-radius: 4px;
                font-size: 10px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s;
                color: white;
                background: rgba(255, 255, 255, 0.1);
            }

            .mode-btn:hover {
                background: rgba(255, 255, 255, 0.2);
                border-color: rgba(255, 255, 255, 0.5);
                transform: translateY(-1px);
            }

            .mode-btn.active {
                background: rgba(255, 255, 255, 0.3);
                border-color: rgba(255, 255, 255, 0.8);
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
            }

            .controls-section {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 6px;
                margin-bottom: 10px;
            }

            .btn {
                padding: 8px 10px;
                border: none;
                border-radius: 5px;
                font-size: 11px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.2s;
                color: white;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
            }

            .btn:disabled {
                opacity: 0.5;
                cursor: not-allowed;
                transform: none !important;
                box-shadow: none !important;
            }

            .btn-primary {
                background: linear-gradient(135deg, #4CAF50, #45a049);
            }

            .btn-primary:hover:not(:disabled) {
                background: linear-gradient(135deg, #45a049, #3d8b40);
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(76, 175, 80, 0.5);
            }

            .btn-primary:active:not(:disabled) {
                transform: translateY(0);
            }

            .btn-danger {
                background: linear-gradient(135deg, #f44336, #d32f2f);
            }

            .btn-danger:hover:not(:disabled) {
                background: linear-gradient(135deg, #d32f2f, #c62828);
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(244, 67, 54, 0.5);
            }

            .btn-danger:active:not(:disabled) {
                transform: translateY(0);
            }

            .btn-secondary {
                background: linear-gradient(135deg, #2196F3, #1976D2);
            }

            .btn-secondary:hover:not(:disabled) {
                background: linear-gradient(135deg, #1976D2, #1565C0);
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(33, 150, 243, 0.5);
            }

            .btn-secondary:active:not(:disabled) {
                transform: translateY(0);
            }

            .info-section {
                background: rgba(255, 255, 255, 0.1);
                border-radius: 4px;
                padding: 8px;
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
                opacity: 0.8;
            }

            .info-value {
                font-weight: 600;
            }

            .kumon-grader-settings {
                border-top: 1px solid rgba(255, 255, 255, 0.2);
                padding: 10px 12px;
                background: rgba(0, 0, 0, 0.2);
            }

            .settings-header {
                font-weight: 600;
                margin-bottom: 8px;
                font-size: 11px;
            }

            .settings-content {
                display: flex;
                flex-direction: column;
                gap: 8px;
            }

            .setting-item {
                font-size: 11px;
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
                border: none;
                border-radius: 4px;
                margin-left: 8px;
                font-size: 12px;
            }

            .setting-item select {
                width: 100px;
                padding: 4px 8px;
                border: none;
                border-radius: 4px;
                margin-left: 8px;
                font-size: 12px;
                background: white;
                color: #333;
            }
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
                e.target.closest('.header-buttons')) {
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
        startButton.addEventListener('click', processAllBoxes);
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

        // Initial refresh
        refreshButton.click();
    }

    // Hotkey handler
    function handleHotkey(event) {
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
        log('Kumon Auto Grader initialized', 'success');
        createUI();
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

