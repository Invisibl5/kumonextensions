// ==UserScript==
// @name         kumonextensions
// @namespace    https://github.com/Invisibl5/kumonextensions
// @version      0.3.10
// @description  Kumon Extensions: Auto Grader + Worksheet Setter
// @author       Invisibl5
// @match        https://class-navi.digital.kumon.com/us/index.html
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
    const CLIENT = {
        applicationName: 'Class-Navi',
        version: '1.0.0.0',
        programName: 'Class-Navi',
        os: typeof navigator !== 'undefined' && navigator.userAgent ? navigator.userAgent : '-',
        machineName: '-'
    };
    const PRESETS = { '5-5': [5, 5], '4-3-3': [4, 3, 3], '3-2-3-2': [3, 2, 3, 2], '2-2-2-2-2': [2, 2, 2, 2, 2], '4-4-2': [4, 4, 2], '3-3-3-1': [3, 3, 3, 1] };

    function INJECT_SCRIPT() {
        const isKumon = (url) => {
            const u = String(url);
            return u.indexOf('digital.kumon.com') !== -1 || u.indexOf('instructor2.') !== -1 || u.indexOf('class-navi.') !== -1;
        };
        const apiName = (url) => { const m = String(url).match(/\/([A-Za-z0-9]+)(?:\?|$)/); return m ? m[1] : url; };
        const isStudyResult = (url) => String(url).indexOf('GetStudyResultInfoList') !== -1;
        const isRegisterStudySet = (url) => String(url).indexOf('RegisterStudySetInfo') !== -1;
        const isList = (url) => {
            const u = String(url);
            return u.indexOf('GetCenterAllStudentList') !== -1 || u.indexOf('StudentList') !== -1 || u.indexOf('GetStudentInfo') !== -1;
        };
        const dispatchToken = (auth) => { if (auth) try { document.dispatchEvent(new CustomEvent('KumonBreakSetToken', { detail: { authorization: auth } })); } catch (e) {} };
        const extractList = (data) => {
            if (!data) return [];
            if (Array.isArray(data)) return data;
            if (data.StudentInfoList && Array.isArray(data.StudentInfoList)) return data.StudentInfoList;
            if (data.CenterAllStudentList && Array.isArray(data.CenterAllStudentList)) return data.CenterAllStudentList;
            if (data.StudentList && Array.isArray(data.StudentList)) return data.StudentList;
            const first = Object.values(data).find(Array.isArray);
            return first || [];
        };
        const safeParse = (s) => { if (!s || typeof s !== 'string') return null; try { return JSON.parse(s); } catch (e) { return null; } };
        const logApi = (dir, name, url, req, res, err) => {
            try {
                document.dispatchEvent(new CustomEvent('KumonBreakSetApiCall', { detail: { dir, apiName: name, url, request: req, response: res, error: err || null } }));
            } catch (e) {}
        };

        const origSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
            if (name === 'Authorization' && value) { this._kumonAuth = value; dispatchToken(value); }
            return origSetRequestHeader.apply(this, arguments);
        };
        const origOpen = XMLHttpRequest.prototype.open;
        const origSend = XMLHttpRequest.prototype.send;
        XMLHttpRequest.prototype.open = function(method, url) { this._kumonSpUrl = url; this._kumonSpMethod = method; return origOpen.apply(this, arguments); };
        XMLHttpRequest.prototype.send = function(body) {
            const xhr = this;
            const url = xhr._kumonSpUrl || '';
            const reqBody = typeof body === 'string' ? body : null;
            const reqParsed = safeParse(reqBody);
            if (isKumon(url)) {
                logApi('XHR_REQ', apiName(url), url, reqParsed || reqBody, null, null);
                xhr.addEventListener('readystatechange', function() {
                    if (xhr.readyState !== 4) return;
                    let resData = null; let err = null;
                    try { resData = xhr.response != null && typeof xhr.response === 'object' ? xhr.response : (xhr.responseText ? JSON.parse(xhr.responseText) : null); } catch (e) { err = e; }
                    logApi('XHR_RES', apiName(url), url, reqParsed || reqBody, resData, err);
                    if (resData && typeof resData === 'object') {
                        if (isStudyResult(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetStudyResult', { detail: { requestBody: reqBody, responseData: resData } }));
                        if (isRegisterStudySet(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetRegister', { detail: { requestBody: reqBody, authorization: xhr._kumonAuth } }));
                        if (isList(url)) { const list = extractList(resData); if (list.length) document.dispatchEvent(new CustomEvent('KumonBreakSetStudents', { detail: { studentsJson: JSON.stringify(list) } })); }
                    }
                });
            }
            return origSend.apply(this, arguments);
        };

        const origFetch = window.fetch;
        window.fetch = function(input, init) {
            const url = typeof input === 'string' ? input : (input && input.url) || '';
            const reqBody = (init && init.body) ? (typeof init.body === 'string' ? init.body : null) : null;
            const reqParsed = safeParse(reqBody);
            let authHeader = (init && init.headers && typeof init.headers.get === 'function' ? init.headers.get('Authorization') : (init.headers && init.headers.Authorization)) || null;
            if (authHeader) dispatchToken(authHeader);
            if (isKumon(url)) logApi('FETCH_REQ', apiName(url), url, reqParsed || reqBody, null, null);
            return origFetch.apply(this, arguments).then(function(res) {
                if (isKumon(url)) {
                    const clone = res.clone();
                    clone.json().then(function(data) {
                        logApi('FETCH_RES', apiName(url), url, reqParsed || reqBody, data, null);
                        if (data && typeof data === 'object') {
                            if (isStudyResult(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetStudyResult', { detail: { requestBody: reqBody, responseData: data } }));
                            if (isRegisterStudySet(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetRegister', { detail: { requestBody: reqBody, authorization: authHeader } }));
                            if (isList(url)) { const list = extractList(data); if (list.length) document.dispatchEvent(new CustomEvent('KumonBreakSetStudents', { detail: { studentsJson: JSON.stringify(list) } })); }
                        }
                    }).catch(function(e) { logApi('FETCH_RES', apiName(url), url, reqParsed || reqBody, null, e); });
                }
                return res;
            });
        };
    }

    function injectBreakSetCapture() {
        if (window.__kumonBreakSetInjected) return;
        const script = document.createElement('script');
        script.textContent = '(' + INJECT_SCRIPT.toString() + ')();';
        const target = document.documentElement || document.head || document.body;
        if (target) {
            window.__kumonBreakSetInjected = true;
            target.appendChild(script);
            script.remove();
        } else {
            document.addEventListener('DOMContentLoaded', function runOnce() {
                document.removeEventListener('DOMContentLoaded', runOnce);
                injectBreakSetCapture();
            });
        }
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
        if (baseSum <= 0 || totalPages % baseSum !== 0) return null;
        const repeats = totalPages / baseSum;
        const expanded = [];
        for (let r = 0; r < repeats; r++) for (let i = 0; i < patternSizes.length; i++) expanded.push(patternSizes[i]);
        return expanded;
    }
    /** Build InsertSetInfoList - multiple break items, ALL with the SAME StudyScheduleIndex. */
    function buildInsertSetInfoList(startPage, totalPages, patternSizes, nextStudyScheduleIndex) {
        const sum = patternSizes.reduce((a, b) => a + b, 0);
        if (sum !== totalPages) return null;
        const index = nextStudyScheduleIndex != null ? nextStudyScheduleIndex : 1;
        const list = []; let from = startPage;
        for (let i = 0; i < patternSizes.length; i++) {
            const n = patternSizes[i];
            const to = from + n - 1;
            list.push({ StudyScheduleIndex: index, WorksheetNOFrom: from, WorksheetNOTo: to, GradingMethod: '1' });
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
                <h3>📌 Kumon Extensions</h3>
                <div class="header-buttons">
                    <button class="collapse-panel" id="kumon-collapse" title="Collapse">▾</button>
                    <select id="kumon-func-select" class="kumon-func-select" title="Switch functionality">
                        <option value="grader">Auto Grader</option>
                        <option value="worksheet-setter">Worksheet Setter</option>
                    </select>
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
            <div id="kumon-panel-worksheet-setter" class="kumon-panel" style="display:none;">
                <div class="kumon-bs-body">
                    <div class="kumon-bs-block">
                        <div class="kumon-bs-label">Current target</div>
                        <div id="kumon-bs-ctx" class="kumon-bs-ctx">(Open a student\u2019s Set page first)</div>
                        <div id="kumon-bs-status" class="kumon-bs-hint">Token \u2717 | Context \u2717</div>
                    </div>
                    <div class="kumon-bs-block">
                        <div class="kumon-bs-label">Start page</div>
                        <input id="kumon-bs-start" type="number" min="1" placeholder="e.g. 111" class="kumon-bs-input" />
                        <div class="kumon-bs-label" style="margin-top:8px;">Total pages</div>
                        <input id="kumon-bs-total" type="number" min="1" placeholder="e.g. 10" class="kumon-bs-input" />
                        <div class="kumon-bs-label" style="margin-top:8px;">Pattern (preset or custom)</div>
                        <div class="kumon-bs-presets">
                            <button type="button" class="kumon-bs-preset" data-pattern="5-5">5-5</button>
                            <button type="button" class="kumon-bs-preset" data-pattern="4-3-3">4-3-3</button>
                            <button type="button" class="kumon-bs-preset" data-pattern="3-2-3-2">3-2-3-2</button>
                            <button type="button" class="kumon-bs-preset" data-pattern="2-2-2-2-2">2-2-2-2-2</button>
                            <button type="button" class="kumon-bs-preset" data-pattern="4-4-2">4-4-2</button>
                            <button type="button" class="kumon-bs-preset" data-pattern="3-3-3-1">3-3-3-1</button>
                        </div>
                        <input id="kumon-bs-pattern" type="text" placeholder="e.g. 5-5 or 5,3,2" class="kumon-bs-input" />
                        <div id="kumon-bs-preview" class="kumon-bs-preview"></div>
                        <button id="kumon-bs-assign-btn" class="kumon-bs-btn">Assign</button>
                        <div id="kumon-bs-result" class="kumon-bs-result"></div>
                        <details class="kumon-bs-details"><summary>Paste context (if capture fails)</summary>
                        <div class="kumon-bs-hint" style="margin:6px 0;">Copy RegisterStudySetInfo request from DevTools \u2192 Network, then paste below.</div>
                        <textarea id="kumon-bs-paste-ctx" class="kumon-bs-textarea" placeholder="Paste full request JSON here"></textarea>
                        <button id="kumon-bs-use-pasted-btn" class="kumon-bs-btn-secondary">Use pasted context</button>
                        </details>
                        <details class="kumon-bs-details"><summary>RegisterStudySetInfo debug</summary>
                        <div class="kumon-bs-hint" style="margin:6px 0;">In console: <code>window.__kumonBreakSetRegisterLog</code> (last 20 calls, <code>source</code> = "page" or "script").</div>
                        <button type="button" id="kumon-bs-copy-last-page" class="kumon-bs-btn-secondary" style="margin-top:6px;">Copy last PAGE request (manual assign)</button>
                        <button type="button" id="kumon-bs-copy-last-script" class="kumon-bs-btn-secondary" style="margin-top:4px;">Copy last SCRIPT request (our Assign)</button>
                        <div id="kumon-bs-register-log-hint" class="kumon-bs-hint" style="margin-top:6px;"></div>
                        </details>
                    </div>
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

            #kumon-extensions-ui:hover {
                transform: translateY(-2px);
                box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4);
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
                flex: 0 0 auto;
            }

            .header-buttons {
                display: flex;
                gap: 6px;
                align-items: center;
            }

            .kumon-func-select {
                padding: 4px 6px;
                font-size: 11px;
                border-radius: 4px;
                border: 1px solid rgba(255,255,255,0.4);
                background: rgba(255,255,255,0.2);
                color: white;
                cursor: pointer;
                pointer-events: auto;
            }
            .kumon-func-select option { background: #333; color: #fff; }
            .kumon-panel { overflow: auto; }

            .toggle-settings, .resize-handle, .collapse-panel {
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

            .toggle-settings:hover, .resize-handle:hover, .collapse-panel:hover {
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
            .kumon-bs-body { padding: 12px; }
            .kumon-bs-block { margin-bottom: 12px; }
            .kumon-bs-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #a6e3a1; margin-bottom: 4px; }
            .kumon-bs-ctx { font-size: 12px; color: #cdd6f4; line-height: 1.4; margin-bottom: 4px; }
            .kumon-bs-hint { font-size: 10px; color: rgba(255,255,255,0.6); }
            .kumon-bs-input { width: 100%; padding: 8px 10px; font-size: 13px; background: rgba(0,0,0,0.35); color: #fff; border: 1px solid rgba(255,255,255,0.3); border-radius: 8px; margin-bottom: 4px; box-sizing: border-box; }
            .kumon-bs-presets { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
            .kumon-bs-preset { padding: 6px 10px; font-size: 11px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.35); background: rgba(0,0,0,0.3); color: #a6e3a1; cursor: pointer; }
            .kumon-bs-preset:hover { background: rgba(255,255,255,0.15); }
            .kumon-bs-preview { font-size: 11px; font-family: monospace; background: rgba(0,0,0,0.35); padding: 8px; border-radius: 6px; margin: 8px 0; color: #a6e3a1; white-space: pre-wrap; word-break: break-all; min-height: 1.2em; }
            .kumon-bs-btn { width: 100%; padding: 10px 14px; background: linear-gradient(135deg, #4CAF50, #45a049); color: #fff; border: none; border-radius: 10px; font-weight: 600; font-size: 13px; cursor: pointer; }
            .kumon-bs-btn:hover { filter: brightness(1.1); }
            .kumon-bs-result { font-size: 12px; margin-top: 8px; min-height: 1.4em; }
            .kumon-bs-details { margin-top: 10px; border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; overflow: hidden; }
            .kumon-bs-details summary { padding: 6px 10px; font-size: 11px; font-weight: 600; cursor: pointer; background: rgba(0,0,0,0.2); }
            .kumon-bs-textarea { width: 100%; min-height: 80px; margin: 6px 0; padding: 8px; font-size: 10px; font-family: monospace; background: rgba(0,0,0,0.35); color: #fff; border: 1px solid rgba(255,255,255,0.2); border-radius: 6px; resize: vertical; box-sizing: border-box; }
            .kumon-bs-btn-secondary { width: 100%; padding: 8px; font-size: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.35); background: rgba(0,0,0,0.3); color: #a6e3a1; cursor: pointer; }
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
                e.target.id === 'kumon-func-select' ||
                (e.target.closest && e.target.closest('#kumon-func-select')) ||
                (e.target.closest && e.target.closest('.header-buttons'))) {
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

        // Functionality dropdown: switch panel content
        const funcSelect = document.getElementById('kumon-func-select');
        const panelGrader = document.getElementById('kumon-panel-grader');
        const panelWorksheetSetter = document.getElementById('kumon-panel-worksheet-setter');
        funcSelect.addEventListener('change', () => {
            const v = funcSelect.value;
            panelGrader.style.display = v === 'grader' ? 'block' : 'none';
            panelWorksheetSetter.style.display = v === 'worksheet-setter' ? 'block' : 'none';
            if (v === 'worksheet-setter') updateWorksheetSetterUI();
        });

        // Worksheet Setter: preview and Assign
        const wsStart = document.getElementById('kumon-bs-start');
        const wsTotal = document.getElementById('kumon-bs-total');
        const wsPattern = document.getElementById('kumon-bs-pattern');
        const wsPreview = document.getElementById('kumon-bs-preview');

        function refreshWsPreview() {
            const start = parseInt(wsStart.value, 10);
            const total = parseInt(wsTotal.value, 10);
            const raw = (wsPattern.value || '').trim();
            const pattern = PRESETS[raw] || parsePattern(raw);
            if (!pattern.length || isNaN(start) || isNaN(total) || start < 1 || total < 1) {
                wsPreview.textContent = '';
                return;
            }
            const expanded = expandPatternToTotal(pattern, total);
            if (!expanded) {
                const baseSum = pattern.reduce((a, b) => a + b, 0);
                wsPreview.textContent = 'Total (' + total + ') must be divisible by pattern sum (' + baseSum + ')';
                wsPreview.style.color = '#f7768e';
                return;
            }
            const ctx = getEffectiveContext();
            const nextIndex = getNextStudyScheduleIndex(ctx, start, total);
            const list = buildInsertSetInfoList(start, total, expanded, nextIndex);
            if (!list) { wsPreview.textContent = ''; return; }
            wsPreview.textContent = list.map(s => 'Set ' + s.StudyScheduleIndex + ': p.' + s.WorksheetNOFrom + '\u2013' + s.WorksheetNOTo).join('\n');
            wsPreview.style.color = '#a6e3a1';
        }

        if (wsPattern) { wsPattern.addEventListener('input', refreshWsPreview); wsPattern.addEventListener('change', refreshWsPreview); }
        if (wsStart) { wsStart.addEventListener('input', refreshWsPreview); wsStart.addEventListener('change', refreshWsPreview); }
        if (wsTotal) { wsTotal.addEventListener('input', refreshWsPreview); wsTotal.addEventListener('change', refreshWsPreview); }

        panelWorksheetSetter.querySelectorAll('.kumon-bs-preset').forEach(btn => {
            btn.addEventListener('click', () => {
                const p = btn.getAttribute('data-pattern');
                if (wsPattern) { wsPattern.value = p; refreshWsPreview(); }
            });
        });

        document.getElementById('kumon-bs-use-pasted-btn').addEventListener('click', function() {
            const raw = (document.getElementById('kumon-bs-paste-ctx').value || '').trim();
            const resultEl = document.getElementById('kumon-bs-result');
            if (!raw) {
                resultEl.textContent = 'Paste the RegisterStudySetInfo request JSON first.';
                resultEl.style.color = '#f7768e';
                return;
            }
            try {
                const parsed = JSON.parse(raw);
                if (parsed && (parsed.StudentID || parsed.CenterID)) {
                    lastRegisterStudySetRequest = parsed;
                    updateWorksheetSetterUI();
                    resultEl.textContent = 'Context applied. You can Assign now.';
                    resultEl.style.color = '#a6e3a1';
                } else {
                    resultEl.textContent = 'JSON must include StudentID and CenterID.';
                    resultEl.style.color = '#f7768e';
                }
            } catch (e) {
                resultEl.textContent = 'Invalid JSON: ' + (e && e.message);
                resultEl.style.color = '#f7768e';
            }
        });

        document.getElementById('kumon-bs-assign-btn').addEventListener('click', function() {
            const resultEl = document.getElementById('kumon-bs-result');
            resultEl.textContent = '';
            resultEl.style.color = '';

            const ctx = getEffectiveContext();
            if (!ctx) {
                resultEl.textContent = 'No context. Open the student\u2019s Set page first.';
                resultEl.style.color = '#f7768e';
                return;
            }
            if (!lastToken) {
                resultEl.textContent = 'No token. Use the app (navigate, open student) so we capture auth.';
                resultEl.style.color = '#f7768e';
                return;
            }

            const start = parseInt(wsStart.value, 10);
            const total = parseInt(wsTotal.value, 10);
            const raw = (wsPattern.value || '').trim();
            const pattern = PRESETS[raw] || parsePattern(raw);

            if (isNaN(start) || start < 1) { resultEl.textContent = 'Enter a valid start page.'; resultEl.style.color = '#f7768e'; return; }
            if (isNaN(total) || total < 1) { resultEl.textContent = 'Enter a valid total pages.'; resultEl.style.color = '#f7768e'; return; }
            if (!pattern.length) { resultEl.textContent = 'Enter a pattern (e.g. 5-5 or 4,3,3).'; resultEl.style.color = '#f7768e'; return; }
            const expanded = expandPatternToTotal(pattern, total);
            if (!expanded) {
                const baseSum = pattern.reduce((a, b) => a + b, 0);
                resultEl.textContent = 'Total (' + total + ') must be divisible by pattern sum (' + baseSum + ')';
                resultEl.style.color = '#f7768e';
                return;
            }

            const nextIndex = getNextStudyScheduleIndex(ctx, start, total);
            const insertList = buildInsertSetInfoList(start, total, expanded, nextIndex);
            if (!insertList) { resultEl.textContent = 'Could not build sets.'; resultEl.style.color = '#f7768e'; return; }

            const payload = {};
            for (const k in ctx) if (ctx.hasOwnProperty(k)) payload[k] = ctx[k];
            payload.InsertSetInfoList = insertList;
            payload.FinishTestSetInfoList = [];
            payload.DeleteSetInfoList = [];
            const nextId = lastApiResponseId != null ? lastApiResponseId + 1 : (ctx.id != null ? parseInt(ctx.id, 10) + 1 : null);
            payload.id = nextId != null ? String(nextId) : String(Date.now());
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

            resultEl.textContent = 'Sending...';

            fetch(REGISTER_STUDY_SET_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Authorization': lastToken },
                body: JSON.stringify(payload)
            }).then(res => res.json().then(data => {
                const rid = data && (data.ID != null ? data.ID : data.id);
                if (rid != null) { const n = parseInt(rid, 10); if (!isNaN(n)) lastApiResponseId = n; }
                registerLog.push({ ts: Date.now(), source: 'script', request: payload, response: data });
                if (registerLog.length > REGISTER_LOG_MAX) registerLog.shift();
                updateWorksheetSetterUI();
                if (res.ok && data.Result && data.Result.ResultCode === 0) {
                    resultEl.textContent = 'Success. ' + insertList.length + ' set(s) assigned.';
                    resultEl.style.color = '#a6e3a1';
                } else {
                    const err = (data.Result && data.Result.Errors && data.Result.Errors[0]) || ('ResultCode=' + (data.Result && data.Result.ResultCode)) || ('status=' + res.status);
                    resultEl.textContent = 'Error: ' + err;
                    resultEl.style.color = '#f7768e';
                }
            }).catch(() => { resultEl.textContent = 'Response error: ' + res.status; resultEl.style.color = '#f7768e'; })).catch(err => {
                resultEl.textContent = 'Request failed: ' + (err && err.message);
                resultEl.style.color = '#f7768e';
            });
        });

        function copyLastBySource(source) {
            for (let i = registerLog.length - 1; i >= 0; i--) {
                if (registerLog[i].source === source && registerLog[i].request) {
                    const json = JSON.stringify(registerLog[i].request, null, 2);
                    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(json);
                    return;
                }
            }
        }
        document.getElementById('kumon-bs-copy-last-page').addEventListener('click', function() { copyLastBySource('page'); updateWorksheetSetterUI(); });
        document.getElementById('kumon-bs-copy-last-script').addEventListener('click', function() { copyLastBySource('script'); updateWorksheetSetterUI(); });

        // Initial refresh
        refreshButton.click();
    }

    function injectWorksheetPerStudyOptions() {
        if (window.__kumonWorksheetPerStudyHooked) return;
        window.__kumonWorksheetPerStudyHooked = true;

        const LABELS = [
            '4-3-3 worksheets per study',
            '3-2 worksheets per study',
            '2-2 worksheets per study'
        ];
        const KEYS = ['4-3-3', '3-2', '2-2'];

        const extendAll = () => {
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
        injectBreakSetCapture();
        log('Kumon Extensions initialized', 'success');
        createUI();
        injectWorksheetPerStudyOptions();
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

