(function() {
    'use strict';
    if (window.__kumonInjectLoaded) return;
    window.__kumonInjectLoaded = true;

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

    function needReplaceAndRequest(url, reqBody) {
        const body = document.body;
        if (!body || body.dataset.kumonExtensionEnabled === 'false') return { need: false };
        const pattern = (body.dataset.kumonWorksheetPattern || '').trim();
        if (!pattern) return { need: false };
        if (!isRegisterStudySet(url)) return { need: false };
        const reqParsed = safeParse(reqBody);
        if (!reqParsed) return { need: false };
        return { need: true, request: reqParsed, patternKey: pattern };
    }

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
        let reqBody = typeof body === 'string' ? body : null;
        const check = needReplaceAndRequest(url, reqBody);
        if (check.need) {
            const handler = function(ev) {
                document.removeEventListener('KumonPayloadReady', handler);
                const payload = ev.detail && ev.detail.payload;
                const newBody = payload ? JSON.stringify(payload) : body;
                origSend.call(xhr, newBody);
            };
            document.addEventListener('KumonPayloadReady', handler);
            try { document.dispatchEvent(new CustomEvent('KumonBuildReplacementPayload', { detail: { request: check.request, patternKey: check.patternKey } })); } catch (e) {}
            return;
        }
        let reqParsed = safeParse(reqBody);
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
        let authHeader = (init && init.headers && typeof init.headers.get === 'function' ? init.headers.get('Authorization') : (init.headers && init.headers.Authorization)) || null;
        if (authHeader) dispatchToken(authHeader);
        const check = needReplaceAndRequest(url, reqBody);
        if (check.need) {
            return new Promise(function(resolve, reject) {
                const handler = function(ev) {
                    document.removeEventListener('KumonPayloadReady', handler);
                    const payload = ev.detail && ev.detail.payload;
                    const initToUse = payload
                        ? Object.assign({}, init, { body: JSON.stringify(payload) })
                        : init;
                    const bodySent = (initToUse && initToUse.body && typeof initToUse.body === 'string') ? initToUse.body : reqBody;
                    const parsedSent = safeParse(bodySent);
                    if (isKumon(url)) logApi('FETCH_REQ', apiName(url), url, parsedSent || bodySent, null, null);
                    origFetch.call(window, input, initToUse).then(function(res) {
                        if (isKumon(url)) {
                            const clone = res.clone();
                            clone.json().then(function(data) {
                                logApi('FETCH_RES', apiName(url), url, parsedSent || bodySent, data, null);
                                if (data && typeof data === 'object') {
                                    if (isStudyResult(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetStudyResult', { detail: { requestBody: bodySent, responseData: data } }));
                                    if (isRegisterStudySet(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetRegister', { detail: { requestBody: bodySent, authorization: authHeader } }));
                                    if (isList(url)) { const list = extractList(data); if (list.length) document.dispatchEvent(new CustomEvent('KumonBreakSetStudents', { detail: { studentsJson: JSON.stringify(list) } })); }
                                }
                            }).catch(function(e) { logApi('FETCH_RES', apiName(url), url, parsedSent || bodySent, null, e); });
                        }
                        resolve(res);
                    }).catch(reject);
                };
                document.addEventListener('KumonPayloadReady', handler);
                try { document.dispatchEvent(new CustomEvent('KumonBuildReplacementPayload', { detail: { request: check.request, patternKey: check.patternKey } })); } catch (e) {
                    origFetch.call(window, input, init).then(resolve).catch(reject);
                }
            });
        }
        const parsedSent = safeParse(reqBody);
        if (isKumon(url)) logApi('FETCH_REQ', apiName(url), url, parsedSent || reqBody, null, null);
        return origFetch.call(this, input, init).then(function(res) {
            if (isKumon(url)) {
                const clone = res.clone();
                clone.json().then(function(data) {
                    logApi('FETCH_RES', apiName(url), url, parsedSent || reqBody, data, null);
                    if (data && typeof data === 'object') {
                        if (isStudyResult(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetStudyResult', { detail: { requestBody: reqBody, responseData: data } }));
                        if (isRegisterStudySet(url)) document.dispatchEvent(new CustomEvent('KumonBreakSetRegister', { detail: { requestBody: reqBody, authorization: authHeader } }));
                        if (isList(url)) { const list = extractList(data); if (list.length) document.dispatchEvent(new CustomEvent('KumonBreakSetStudents', { detail: { studentsJson: JSON.stringify(list) } })); }
                    }
                }).catch(function(e) { logApi('FETCH_RES', apiName(url), url, parsedSent || reqBody, null, e); });
            }
            return res;
        });
    };
})();
