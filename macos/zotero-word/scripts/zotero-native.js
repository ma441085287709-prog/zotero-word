/* Runs one accepted job inside Zotero. MacWord supplies fresh field handles for each document. */
async function runJob(jobDir, sessionID, stopped) {
    const integration = Zotero.Integration;
    const path = name => jobDir + "/" + name;
    const pause = () => Zotero.Promise.delay(100);
    const check = (ok, message) => { if (!ok) throw new Error(message); };
    const read = async name => JSON.parse(await Zotero.File.getContentsAsync(path(name)));
    const exists = name => Zotero.File.pathToFile(path(name)).exists();
    const write = (name, value) => Zotero.File.putContentsAsync(path(name), JSON.stringify(value));
    const checkAbort = () => check(!stopped.exists() && !exists("abort"), "Word controller stopped; inspect the working copy");
    let stage = "connect";
    let refreshTiming;
    let timingWrite = Promise.resolve();
    async function signal(name, value = {}) {
        stage = name;
        checkAbort();
        await write(name + ".json", value);
    }
    async function waitFor(label, predicate, ms = 30000) {
        const until = Date.now() + ms;
        while (Date.now() < until) {
            checkAbort();
            const value = predicate();
            if (value) return value;
            await pause();
        }
        throw new Error("Timed out waiting for " + label);
    }
    async function handshake(name, value) {
        await signal(name, value);
        await waitFor(name, () => exists(name + ".ok"));
    }
    const claim = Zotero.File.pathToFile(path("dispatched"));
    try {
        claim.create(0, 0o600);
    } catch (error) {
        if (claim.exists()) return false;
        throw error;
    }
    try {
        const cfg = await read("job.json");
        const operation = cfg.libraryOnly ? (cfg.imports?.length ? "library_import" : "library_read")
            : (cfg.reopenOnly ? "reopen_verify" : "refresh");
        if (operation === "refresh") {
            refreshTiming = { scriptVersion: "2.4.3-slow-update-fix.2", setDocPrefsRequired: null,
                styleConfirmedAt: null, setDocPrefsReturnedAt: null, fieldsReadAt: null, writeErrors: [] };
        }
        const accepted = { job: jobDir.slice(jobDir.lastIndexOf("/") + 1), session: sessionID,
            operation, document: cfg.document || null };
        const request = await read("request.json");
        check(Object.keys(accepted).every(key => request[key] === accepted[key]),
            "The native request belongs to another job, session or document");
        await signal("started", accepted);
        if (cfg.libraryOnly) {
            const lib = Zotero.Libraries.userLibraryID;
            const library_id = Zotero.URI.getLibraryPath(lib);
            check(!cfg.libraryID || cfg.libraryID === library_id,
                "The requested library differs from the open personal library. Supply a current snapshot of the requested library.");
            const receipt = [];
            if (cfg.imports?.length) {
                check(cfg.libraryID === library_id, "Import requires the bound personal library");
                stage = "import missing references";
                const normalize = doi => String(doi || "").trim().toLowerCase().replace(/^(?:doi:\s*|https?:\/\/(?:dx\.)?doi\.org\/)/, "");
                await Zotero.DB.executeTransaction(async () => {
                    const current = (await Zotero.Items.getAll(lib, true, false)).filter(i => !i.deleted && i.isRegularItem());
                    await Zotero.Items.loadDataTypes(current, ["itemData"]);
                    const byDOI = new Map();
                    for (const item of current) {
                        const extraDOI = /^\s*DOI\s*:\s*(\S+)/im.exec(item.getField("extra") || "");
                        const doi = normalize(item.getField("DOI") || extraDOI?.[1]);
                        if (doi) byDOI.set(doi, [...(byDOI.get(doi) || []), item]);
                    }
                    for (const data of cfg.imports) {
                        checkAbort();
                        const doi = normalize(data.DOI);
                        check(/^10\.\d{4,9}\/\S+$/.test(doi) && data.title && data.type, "Missing exact DOI metadata");
                        const hits = byDOI.get(doi) || [];
                        if (hits.length > 1) {
                            receipt.push({ doi, status: "ambiguous" });
                            continue;
                        }
                        let item = hits[0];
                        if (!item) {
                            item = new Zotero.Item();
                            item.libraryID = lib;
                            Zotero.Utilities.Item.itemFromCSLJSON(item, data);
                            await item.save();
                            check(normalize(item.getField("DOI") || Zotero.Utilities.Item.itemToCSLJSON(item).DOI) === doi,
                                "Zotero did not retain the imported DOI: " + doi);
                            byDOI.set(doi, [item]);
                        }
                        receipt.push({ doi, uri: Zotero.URI.getItemURI(item), created: !hits.length });
                    }
                    checkAbort();
                });
            }
            const items = [];
            const records = (await Zotero.Items.getAll(lib, true, false)).filter(item => !item.deleted && item.isRegularItem());
            await Zotero.Items.loadDataTypes(records, ["itemData", "creators"]);
            for (const item of records) {
                const csl = Zotero.Utilities.Item.itemToCSLJSON(item);
                csl.id = Zotero.URI.getItemURI(item);
                items.push(csl);
            }
            await signal("library", { library_id, items, import_receipt: receipt, observed_at: new Date().toISOString() });
            return true;
        }
        await Zotero.Styles.init();
        const style = Zotero.Styles.get(cfg.style);
        check(style && style.styleID === cfg.style, "Target CSL style is not installed: " + cfg.style);
        const app = integration.getApplication("MacWord16", "refresh");
        const byURI = new Map();
        async function resolve(aliases) {
            for (const uri of aliases) {
                if (!byURI.has(uri)) byURI.set(uri, await Zotero.URI.getURIItem(uri));
                const item = byURI.get(uri);
                if (item && !item.deleted && item.isRegularItem()) return item;
            }
            throw new Error("No active Zotero item for " + aliases.join(" / "));
        }
        const expected = new Map();
        for (const row of cfg.citations) {
            const items = [];
            for (const aliases of row.items) items.push((await resolve(aliases)).id);
            expected.set(row.id, items);
        }
        function equal(actual, wanted) {
            return actual.length === wanted.length && actual.every((value, i) => value === wanted[i]);
        }
        function visible(text) {
            // The DOCX reader and MacWord expose different paragraph and tab separators.
            return cfg.savedCopy ? text.normalize("NFC").replace(/\s+/gu, "") : text;
        }
        function checkBibliography(bib) {
            check(bib && bib[1].length, "Zotero did not render bibliography entries for this style");
            const ids = new Set(bib[0].entry_ids.flat().map(Number));
            const cited = new Set([...expected.values()].flat());
            check(ids.size === cited.size && [...cited].every(id => ids.has(id)), "Bibliography and cited item sets differ");
            return bib[1];
        }
        async function fields(doc, formatted) {
            const citations = new Map();
            const bibliographies = [];
            for (const field of await doc.getFields("Field")) {
                const code = (await field.getCode()).trim();
                const match = /^(?:ITEM\s+CSL_CITATION|CITATION)\s*(\{[\s\S]+})$/.exec(code);
                if (match) {
                    const data = JSON.parse(match[1]);
                    const wanted = expected.get(data.citationID);
                    check(wanted && !citations.has(data.citationID), "Unexpected or duplicate native citation " + data.citationID);
                    const ids = [];
                    for (const item of data.citationItems) ids.push((await resolve(item.uris)).id);
                    check(equal(ids, wanted), "Different items or order in " + data.citationID);
                    const text = formatted ? String(await field.getText()).trim() : "";
                    check(!formatted || (text && text !== "{Citation}"), "Unformatted citation " + data.citationID);
                    citations.set(data.citationID, { field, text });
                } else if (/^BIBL\b/.test(code)) {
                    bibliographies.push(formatted ? String(await field.getText()).trim() : "");
                } else if (/^(?:ITEM|CITATION|TEMP)\b/.test(code)) {
                    throw new Error("Unrecognized or pending Zotero field");
                }
            }
            check(citations.size === expected.size, "Missing native citation fields");
            check(formatted ? bibliographies.length === 1 && bibliographies[0] : bibliographies.length <= 1,
                "Expected one formatted native bibliography after refresh");
            const result = { citations, bibliographies };
            if (!formatted) result.readAt = new Date().toISOString();
            return result;
        }
        function recordTiming(update) {
            Object.assign(refreshTiming, update);
            const snapshot = { ...refreshTiming };
            // Dialog events and command completion can arrive during the same file write.
            timingWrite = timingWrite.then(() => write("refresh-timing.json", snapshot))
                .catch(error => {
                    refreshTiming.writeErrors.push({ at: new Date().toISOString(),
                        fields: Object.keys(update), error: String(error) });
                    try { Zotero.logError(error); } catch { /* The receipt retains the original error. */ }
                });
            return timingWrite;
        }
        function prefs(docData) {
            const data = new integration.DocumentData(docData);
            return data.style.styleID === cfg.style && data.prefs.fieldType === "Field"
                && (cfg.abbreviations === null || data.prefs.automaticJournalAbbreviations === cfg.abbreviations);
        }
        async function attach(label, action) {
            checkAbort();
            check(!integration.currentDoc, "Another Zotero integration command is active");
            const done = Zotero.Promise.defer();
            integration.currentDoc = true;
            integration.currentCommandPromise = done.promise;
            let doc;
            try {
                integration.currentDoc = doc = await app.getActiveDocument();
                const observed = await fields(doc, label === "reattached");
                const id = (cfg.citations.find(row => row.story === "word/document.xml") || cfg.citations[0]).id;
                // MacWord's getDocument(path) takes an application path. Match a
                // selected live field to the exact file Word has open before writing.
                await observed.citations.get(id).field.select();
                await handshake(label, { citation: id, bibliographyCount: observed.bibliographies.length,
                    scriptVersion: refreshTiming?.scriptVersion });
                const data = await doc.getDocumentData();
                check(data !== "ZOTERO_TRANSFER_DOCUMENT", "An exported document needs an explicit import");
                return await action(doc, observed, data);
            } finally {
                try {
                    if (doc) {
                        try { await doc.cleanup(); } finally { if (doc.complete) await doc.complete(); }
                    }
                } finally {
                    integration.currentDoc = false;
                    done.resolve();
                }
            }
        }
        async function runWithoutSlowUpdatePrompt(api, session, name) {
            const initial = session.data.prefs;
            const original = { ...initial };
            let result;
            try {
                initial.dontAskDelayCitationUpdates = true;
                result = await api[name]();
            } finally {
                // setDocPrefs copies preferences to a new object and can turn a missing value into false.
                for (const target of new Set([initial, session.data?.prefs ?? initial])) {
                    for (const key of ["dontAskDelayCitationUpdates", "delayCitationUpdates"]) {
                        if (Object.hasOwn(original, key)) target[key] = original[key];
                        else delete target[key];
                    }
                }
            }
            if (name === "setDocPrefs") {
                await recordTiming({ setDocPrefsReturnedAt: new Date().toISOString() });
            }
            return result;
        }
        async function command(doc, name, interact) {
            checkAbort();
            const [session, imported] = await integration.getSession(app, doc, "MacWord16", name);
            integration.currentSession = session;
            try {
                check(!imported, "An exported document needs an explicit import");
                const api = new integration.Interface(app, doc, session);
                // execCommand catches plugin errors. Await the same native command
                // body directly so a failed refresh cannot turn into a receipt.
                const previousWindow = integration.currentWindow;
                const work = name === "setDocPrefs" || name === "refresh"
                    ? runWithoutSlowUpdatePrompt(api, session, name) : api[name]();
                if (interact) {
                    const interaction = interact(previousWindow).catch(error => {
                        const win = integration.currentWindow;
                        const file = name === "setDocPrefs" ? "integrationDocPrefs.xhtml" : "citationDialog.xhtml";
                        if (integration.currentDoc === doc && win && win !== previousWindow && !win.closed &&
                                String(win.location.href) === "chrome://zotero/content/integration/" + file) {
                            win.close();
                        }
                        throw error;
                    });
                    const outcomes = await Promise.allSettled([work, interaction]);
                    // The window must close and the native call must unwind before cleanup.
                    for (const outcome of outcomes.reverse()) {
                        if (outcome.status === "rejected") throw outcome.reason;
                    }
                } else {
                    await work;
                }
                checkAbort();
                if (name !== "editCitation") await doc.setDocumentData(session.data.serialize());
                return session;
            } finally {
                if (session.progressBar) await session.progressBar.hide(true);
            }
        }
        async function dialog(file, previousWindow, ready) {
            return waitFor(file, () => {
                const win = integration.currentWindow;
                return win && win !== previousWindow && !win.closed && String(win.location.href) ===
                    "chrome://zotero/content/integration/" + file && ready(win) && win;
            });
        }
        async function preferences(previousWindow) {
            const win = await dialog("integrationDocPrefs.xhtml", previousWindow, w => {
                if (w.document.readyState !== "complete") return false;
                const d = w.document.querySelector("dialog");
                return d?._buttons?.accept && !d.getButton("accept").disabled && w.document.getElementById("style-configurator")?.style;
            });
            const d = win.document;
            const box = d.querySelector("dialog");
            try {
                const selector = d.getElementById("style-configurator");
                selector.style = cfg.style;
                check(selector.style === cfg.style, "Zotero could not select the target style");
                const bookmarks = d.getElementById("formatUsingBookmarks");
                if (bookmarks) bookmarks.checked = false;
                if (cfg.abbreviations !== null) {
                    const abbreviations = d.getElementById("automaticJournalAbbreviations");
                    check(abbreviations && (!abbreviations.disabled || abbreviations.checked === cfg.abbreviations),
                        "The style does not support the requested abbreviation setting");
                    abbreviations.checked = cfg.abbreviations;
                }
                box.addEventListener("dialogclosing", event => {
                    if (event.detail.button === "accept") {
                        void recordTiming({ styleConfirmedAt: new Date().toISOString() });
                    }
                }, { once: true });
                box.acceptDialog();
            } catch (e) {
                box.cancelDialog();
                throw e;
            }
        }
        if (!cfg.reopenOnly) {
            const rendered = await attach("attached", async (doc, current, data) => {
                const prefData = new integration.DocumentData(data);
                if (prefData.sessionID && integration.sessions[prefData.sessionID]) {
                    // A copied document may still name the open original's cached Zotero session.
                    prefData.sessionID = Zotero.Utilities.randomString(8);
                    await doc.setDocumentData(prefData.serialize());
                }
                let observedFields = current;
                if (!prefs(data)) {
                    stage = "preferences";
                    await recordTiming({ setDocPrefsRequired: true });
                    await command(doc, "setDocPrefs", preferences);
                    check(prefs(await doc.getDocumentData()), "Zotero did not retain the requested preferences");
                    observedFields = await fields(doc, false);
                    await recordTiming({ fieldsReadAt: observedFields.readAt });
                } else {
                    await recordTiming({ setDocPrefsRequired: false, styleConfirmedAt: new Date().toISOString(),
                        fieldsReadAt: current.readAt });
                }
                const bibliographyCount = observedFields.bibliographies.length;
                await handshake("refresh-target", { bibliographyCount, refreshTiming });
                if (!bibliographyCount) {
                    await command(doc, "addEditBibliography");
                    await handshake("bibliography-inserted", {});
                }
                stage = "refresh";
                const start = Date.now();
                const session = await command(doc, "refresh");
                const refreshSeconds = (Date.now() - start) / 1000;
                const observed = await fields(doc, true);
                const entries = checkBibliography(session.bibliography.getCiteprocBibliography(session.style));
                return { text: observed.bibliographies[0], entryCount: entries.length, refreshSeconds,
                    citations: [...observed.citations].map(([id, v]) => [id, v.text]) };
            });
            await signal("refreshed", rendered);
            return true;
        }
        const rendered = await read(cfg.savedCopy ? "saved-copy.json" : "refreshed.json");
        let entryCount = rendered.entryCount;
        const renderedIDs = new Set(rendered.citations.map(([id]) => id));
        check(rendered.citations.length === expected.size && renderedIDs.size === expected.size
            && [...expected.keys()].every(id => renderedIDs.has(id)) && (cfg.savedCopy || entryCount > 0),
            "The saved evidence does not cover the planned fields and bibliography");
        const samples = [];
        await attach("reattached", async (doc, saved, data) => {
            check(prefs(data), "Saved style or abbreviation setting changed");
            check(visible(saved.bibliographies[0]) === visible(rendered.text), "Saved bibliography differs from its evidence");
            for (const [id, text] of rendered.citations) {
                check(visible(saved.citations.get(id).text) === visible(text), "Saved citation differs: " + id);
            }
            if (cfg.savedCopy) {
                const [session, imported] = await integration.getSession(app, doc, "MacWord16", "refresh");
                check(!imported, "The saved copy needs an explicit import");
                integration.currentSession = session;
                let processor;
                try {
                    // Render each group in document order, with an engine independent of the Word session.
                    await session.updateFromDocument();
                    const clusters = Object.values(session.citationsByIndex).map(citation => citation.toJSON());
                    processor = style.getCiteProc(session.data.style.locale, "text", {
                        automaticJournalAbbreviations: session.data.prefs.automaticJournalAbbreviations
                    });
                    if (processor.opt.development_extensions) processor.opt.development_extensions.wrap_url_and_doi = false;
                    const rust = typeof processor.getBatchedUpdates === "function";
                    const outputs = processor.rebuildProcessorState(clusters, rust ? "plain" : "text",
                        Array.from(session.bibliography.uncitedItemIDs));
                    const byID = new Map(rust ? processor.getBatchedUpdates().clusters
                        : outputs.map(([id, , text]) => [id, text]));
                    check(clusters.length === expected.size, "Zotero loaded a different number of saved citation groups");
                    for (const { citationID, properties } of clusters) {
                        const formatted = properties.dontUpdate || properties.custom ? properties.plainCitation : byID.get(citationID);
                        check(typeof formatted === "string" && saved.citations.has(citationID) &&
                            visible(formatted) === visible(saved.citations.get(citationID).text),
                            "Saved citation text differs from Zotero: " + citationID);
                    }
                    const entries = checkBibliography(session.bibliography.getCiteprocBibliography(processor));
                    check(visible(entries.join("")) === visible(saved.bibliographies[0]), "Saved bibliography text differs from Zotero");
                    entryCount = entries.length;
                } finally {
                    try { if (processor) processor.free(); }
                    finally { if (session.progressBar) await session.progressBar.hide(true); }
                }
            }
            for (const id of cfg.samples) {
                stage = "editCitation " + id;
                await saved.citations.get(id).field.select();
                let recovered;
                try {
                    await command(doc, "editCitation", async previousWindow => {
                        const win = await dialog("citationDialog.xhtml", previousWindow, w => w.DIALOG_STATE?.loaded);
                        const cancel = win.document.getElementById("cancel-button");
                        check(cancel, "Cannot cancel Zotero's citation editor");
                        try {
                            const io = win.arguments[0].wrappedJSObject;
                            await io.allCitedDataLoadedPromise;
                            check(io.citation.citationID === id, "Zotero opened a different citation");
                            const items = [];
                            const planned = cfg.citations.find(row => row.id === id).items;
                            for (const citationItem of io.citation.citationItems) {
                                const nativeID = Number(citationItem.id);
                                check(Number.isSafeInteger(nativeID) && nativeID > 0, "Zotero restored an embedded item");
                                const item = await Zotero.Items.getAsync(nativeID);
                                check(item && !item.deleted && item.isRegularItem(), "Deleted or missing library item");
                                const matchedURI = planned[items.length]?.find(uri => byURI.get(uri)?.id === item.id);
                                check(matchedURI, "Edit Citation returned an unexpected item in " + id);
                                items.push({ id: item.id, uri: Zotero.URI.getItemURI(item), matchedURI });
                            }
                            check(equal(items.map(i => i.id), expected.get(id)), "Edit Citation restored different item identities");
                            recovered = items;
                        } finally {
                            cancel.click();
                            await pause();
                            if (!win.closed) win.close();
                            await waitFor("the sampled citation editor to close", () => win.closed, 5000);
                        }
                    });
                } catch (e) {
                    if (!recovered || !(e instanceof Zotero.Exception.UserCancelled)) throw e;
                }
                check(recovered, "Edit Citation did not recover the requested items");
                samples.push({ citation: id, items: recovered });
            }
        });
        await signal("verified", { document: cfg.document, sha256: cfg.documentHash, citations: expected.size,
            bibliographyEntries: entryCount, samples, style: cfg.style });
        return true;
    } catch (error) {
        // A dialog event can still have a timing write in flight when the native command fails.
        if (refreshTiming) await timingWrite;
        await write("error.json", { stage, type: error.name, error: String(error), refreshTiming });
        return false;
    }
}
