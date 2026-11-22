/* Runs one accepted job inside Zotero. WinWord supplies fresh field handles for each document. */
var winWordExecutorVersion = "windows-native-selection-2.4.2";
var winWordScriptVersion = "2.4.3-direct-save.2-handoff.1";
async function runJob(jobDir, sessionID, stopped) {
    const integration = Zotero.Integration;
    const path = name => jobDir + "\\" + name;
    const pause = () => Zotero.Promise.delay(100);
    const check = (ok, message) => { if (!ok) throw new Error(message); };
    const read = async name => JSON.parse(await Zotero.File.getContentsAsync(path(name)));
    const exists = name => Zotero.File.pathToFile(path(name)).exists();
    const write = (name, value) => Zotero.File.putContentsAsync(path(name), JSON.stringify(value));
    const checkAbort = () => check(!stopped.exists() && !exists("abort"), "Word controller stopped; inspect the working copy");
    let stage = "connect";
    let refreshTiming;
    let timingWrite = Promise.resolve();
    function recordTiming(update) {
        Object.assign(refreshTiming, update);
        const snapshot = { ...refreshTiming };
        // The ribbon callback and the field reader can finish during the same write.
        timingWrite = timingWrite.then(() => write("refresh-timing.json", snapshot)).catch(error => {
            refreshTiming.writeErrors.push({ at: new Date().toISOString(), fields: Object.keys(update), error: String(error) });
            try { Zotero.logError(error); } catch { /* The main receipt retains the error. */ }
        });
        return timingWrite;
    }
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
            refreshTiming = { scriptVersion: winWordScriptVersion, setDocPrefsRequired: null,
                styleConfirmedAt: null, setDocPrefsReturnedAt: null, fieldsReadAt: null, writeErrors: [] };
        }
        const accepted = { job: Zotero.File.pathToFile(jobDir).leafName, session: sessionID,
            operation, document: cfg.document || null };
        const request = await read("request.json");
        if (request.executor !== undefined) accepted.executor = winWordExecutorVersion;
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
        const app = integration.getApplication("WinWord", "refresh", cfg.document);
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
            // The DOCX reader and WinWord expose different paragraph and tab separators.
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
                    check(formatted !== true || (text && text !== "{Citation}"), "Unformatted citation " + data.citationID);
                    citations.set(data.citationID, { field, text });
                } else if (/^BIBL\b/.test(code)) {
                    bibliographies.push(formatted ? String(await field.getText()).trim() : "");
                } else if (/^(?:ITEM|CITATION|TEMP)\b/.test(code)) {
                    throw new Error("Unrecognized or pending Zotero field");
                }
            }
            check(citations.size === expected.size, "Missing native citation fields");
            check(formatted === true ? bibliographies.length === 1 && bibliographies[0] : bibliographies.length <= 1,
                "Expected one formatted native bibliography after refresh");
            return { citations, bibliographies, readAt: formatted === false ? new Date().toISOString() : undefined };
        }
        function prefs(docData) {
            const data = new integration.DocumentData(docData);
            return data.style?.styleID === cfg.style && data.prefs?.fieldType === "Field"
                && (cfg.abbreviations === null || data.prefs.automaticJournalAbbreviations === cfg.abbreviations);
        }
        async function confirmSelectedField(doc, id) {
            const field = await doc.cursorInField("Field");
            check(field, "Zotero could not see Word's selected field " + id);
            const match = /^(?:ITEM\s+CSL_CITATION|CITATION)\s*(\{[\s\S]+\})$/.exec((await field.getCode()).trim());
            check(match && JSON.parse(match[1]).citationID === id, "Word and Zotero selected different fields");
        }
        async function attach(label, action) {
            checkAbort();
            check(!integration.currentDoc, "Another Zotero integration command is active");
            const done = Zotero.Promise.defer();
            integration.currentDoc = true;
            integration.currentCommandPromise = done.promise;
            let doc;
            try {
                // Word's OLE calls must start outside the original input event.
                await Zotero.Promise.delay();
                integration.currentDoc = doc = await app.getDocument(cfg.document);
                const observed = await fields(doc, label === "reattached");
                if (label === "after-preferences") await recordTiming({ fieldsReadAt: observed.readAt });
                const id = (cfg.citations.find(row => row.story === "word/document.xml") || cfg.citations[0]).id;
                // Word selects its real field and Zotero independently reads it back.
                await handshake(label, { citation: id, bibliographyCount: observed.bibliographies.length,
                    scriptVersion: refreshTiming?.scriptVersion });
                await confirmSelectedField(doc, id);
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
        // WinWord owns the ribbon command. Scope the prompt preference to the active
        // session's actual field update so Zotero restores it before serializing the document.
        function suppressSlowPrompt(session, nativeCommand) {
            const update = session.updateDocument;
            const own = Object.getOwnPropertyDescriptor(session, "updateDocument");
            async function updateWithoutPrompt(...args) {
                if (this !== session || integration.currentCommandPromise !== nativeCommand) return update.apply(this, args);
                const settings = session.data.prefs;
                const previous = Object.getOwnPropertyDescriptor(settings, "dontAskDelayCitationUpdates");
                try {
                    settings.dontAskDelayCitationUpdates = true;
                    return await update.apply(this, args);
                } finally {
                    if (previous) Object.defineProperty(settings, "dontAskDelayCitationUpdates", previous);
                    else delete settings.dontAskDelayCitationUpdates;
                }
            }
            session.updateDocument = updateWithoutPrompt;
            return () => {
                if (session.updateDocument !== updateWithoutPrompt) return;
                if (own) Object.defineProperty(session, "updateDocument", own);
                else delete session.updateDocument;
            };
        }
        async function command(doc, name, interact) {
            checkAbort();
            const [session, imported] = await integration.getSession(app, doc, "WinWord", name);
            integration.currentSession = session;
            let restorePrompt;
            try {
                check(!imported, "An exported document needs an explicit import");
                const api = new integration.Interface(app, doc, session);
                // execCommand catches plugin errors. Await the same native command
                // body directly so a failed refresh cannot turn into a receipt.
                const previousWindow = integration.currentWindow;
                if (name === "refresh") restorePrompt = suppressSlowPrompt(session, integration.currentCommandPromise);
                const work = api[name]();
                if (interact) {
                    const interaction = interact(previousWindow).catch(error => {
                        const win = integration.currentWindow;
                        const file = "citationDialog.xhtml";
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
                restorePrompt?.();
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
        async function configureInWord() {
            // Observe the registered WinWord command before Word invokes its ribbon button.
            const previous = integration.currentCommandPromise;
            const oldWindow = integration.currentWindow;
            const handled = new WeakSet();
            let dispatch = null, commandDone = null, watchedWindow = null, prompted = false, fallbackError = null;
            const started = Date.now();
            await recordTiming({ setDocPrefsRequired: true });
            await signal("settings", { needed: true, document: cfg.document, style: cfg.style });
            async function prompt() {
                if (prompted) return;
                prompted = true;
                await write("preferences-user.json", { document: cfg.document, style: cfg.style,
                    message: "Confirm the target style and Word Fields in the open Zotero Document Preferences window." });
            }
            while (Date.now() - started < 180000) {
                checkAbort();
                if (!dispatch && exists("settings.ok")) {
                    dispatch = await read("settings.ok");
                    check(["invoked", "unknown", "unavailable"].includes(dispatch), "Unknown Word preference dispatch result");
                    if (dispatch === "unavailable") {
                        check(!integration.currentDoc, "Another integration command owns the desktop");
                        // The same named WinWord command, explicitly bound to this document.
                        void integration.execCommand("WinWord", "setDocPrefs", cfg.document, 1).catch(e => { fallbackError = e; });
                    }
                }
                if (fallbackError) throw fallbackError;
                if (!commandDone && integration.currentCommandPromise && integration.currentCommandPromise !== previous) {
                    commandDone = integration.currentCommandPromise;
                    // Word may acknowledge the ribbon before Zotero finishes the native command.
                    void commandDone.then(() => {
                        if (watchedWindow) return recordTiming({ setDocPrefsReturnedAt: new Date().toISOString() });
                    }, error => { fallbackError = error; });
                }
                const win = integration.currentWindow;
                if (commandDone && integration.currentCommandPromise === commandDone && integration.currentDoc &&
                        win && win !== oldWindow && !win.closed && !handled.has(win)) {
                    const d = win.document;
                    const selector = d.getElementById("style-configurator");
                    if (integration.currentWindowType === "documentPreferences" || selector) {
                        if (win !== watchedWindow) {
                            watchedWindow = win;
                            const session = integration.currentSession;
                            const doc = integration.currentDoc;
                            win.addEventListener("dialogclosing", event => {
                                if (event.detail?.button !== "accept" || integration.currentCommandPromise !== commandDone ||
                                        integration.currentDoc !== doc || session?._doc !== doc || session.agent !== "WinWord" ||
                                        stopped.exists() || exists("abort")) return;
                                const io = win.arguments?.[0]?.wrappedJSObject;
                                // The dialog converts an unset automatic-update flag into a boolean.
                                if (io) io.delayCitationUpdates = session.data.prefs.delayCitationUpdates;
                                const restore = suppressSlowPrompt(session, commandDone);
                                void commandDone.then(restore, restore);
                                if (!refreshTiming.styleConfirmedAt && (io?.style ?? d.getElementById("style-configurator")?.style) === cfg.style) {
                                    void recordTiming({ styleConfirmedAt: new Date().toISOString() });
                                }
                            }, { once: true });
                        }
                        const box = d.querySelector("dialog");
                        let accept;
                        if (d.readyState === "complete") {
                            try { accept = box?.getButton?.("accept"); } catch { /* The dialog is still initializing. */ }
                        }
                        if (typeof selector?.style === "string" && accept && !accept.disabled) {
                            try {
                                if (selector.style !== cfg.style) {
                                    selector.style = cfg.style;
                                    await pause();
                                }
                                const bookmarks = d.getElementById("formatUsingBookmarks");
                                const abbrev = d.getElementById("automaticJournalAbbreviations");
                                if (bookmarks) bookmarks.checked = false;
                                if (cfg.abbreviations !== null && abbrev && !abbrev.disabled) abbrev.checked = cfg.abbreviations;
                                handled.add(win);
                                if (selector.style === cfg.style && (!bookmarks || !bookmarks.checked) &&
                                        (cfg.abbreviations === null || abbrev?.checked === cfg.abbreviations)) {
                                    if (box.acceptDialog() === false) {
                                        await prompt();
                                    } else if (!refreshTiming.styleConfirmedAt) {
                                        await recordTiming({ styleConfirmedAt: new Date().toISOString() });
                                    }
                                } else {
                                    await prompt();
                                }
                            } catch (error) {
                                if (!win.closed) { handled.add(win); await prompt(); }
                            }
                        } else if (Date.now() - started > 2000) {
                            await prompt();
                        }
                    }
                }
                if (dispatch && commandDone && !integration.currentDoc) {
                    await commandDone;
                    if (!refreshTiming.setDocPrefsReturnedAt) {
                        await recordTiming({ setDocPrefsReturnedAt: new Date().toISOString() });
                    }
                    return;
                }
                await pause();
            }
            throw new Error("Document Preferences is unresolved; preserve this working copy and resume after confirming it");
        }
        async function compareStyle(doc, saved) {
            const display = text => text.normalize("NFC").replace(/\s+/gu, "");
            const [session, imported] = await integration.getSession(app, doc, "WinWord", "refresh");
            check(!imported, "The document needs an explicit import");
            integration.currentSession = session;
            let processor;
            try {
                // Render each group in document order, with an engine independent of the Word session.
                await session.updateFromDocument();
                const clusters = Object.values(session.citationsByIndex).map(citation => citation.toJSON());
                processor = style.getCiteProc(session.data.style.locale, "text", {
                    automaticJournalAbbreviations: session.data.prefs.automaticJournalAbbreviations
                });
                if (processor.opt?.development_extensions) processor.opt.development_extensions.wrap_url_and_doi = false;
                const rust = typeof processor.getBatchedUpdates === "function";
                const outputs = processor.rebuildProcessorState(clusters, rust ? "plain" : "text",
                    Array.from(session.bibliography.uncitedItemIDs));
                const byID = new Map(rust ? processor.getBatchedUpdates().clusters
                    : outputs.map(([id, , text]) => [id, text]));
                check(clusters.length === expected.size, "Zotero loaded a different number of saved citation groups");
                for (const { citationID, properties } of clusters) {
                    const formatted = properties.dontUpdate || properties.custom ? properties.plainCitation : byID.get(citationID);
                    if (typeof formatted !== "string" || !saved.citations.has(citationID) ||
                            display(formatted) !== display(saved.citations.get(citationID).text)) {
                        return {difference: "Citation text differs from Zotero: " + citationID};
                    }
                }
                const entries = checkBibliography(session.bibliography.getCiteprocBibliography(processor));
                if (display(entries.join("")) !== display(saved.bibliographies[0])) {
                    return {difference: "Bibliography text differs from Zotero"};
                }
                return {entryCount: entries.length};
            } finally {
                try { if (processor) processor.free(); }
                finally { if (session.progressBar) await session.progressBar.hide(true); }
            }
        }
        if (!cfg.reopenOnly) {
            async function refresh(doc, current) {
                const bibliographyCount = current.bibliographies.length;
                await handshake("refresh-target", { bibliographyCount, refreshTiming });
                if (!bibliographyCount) {
                    await command(doc, "addEditBibliography");
                    await handshake("bibliography-inserted", {});
                }
                stage = "refresh";
                const start = Date.now();
                await signal("refresh-issued", { document: cfg.document });
                const session = await command(doc, "refresh");
                const refreshSeconds = (Date.now() - start) / 1000;
                const observed = await fields(doc, true);
                const entries = checkBibliography(session.bibliography.getCiteprocBibliography(session.style));
                return { text: observed.bibliographies[0], entryCount: entries.length, refreshSeconds,
                    citations: [...observed.citations].map(([id, v]) => [id, v.text]) };
            }
            let rendered = await attach("attached", async (doc, current, data) => {
                const prefData = new integration.DocumentData(data);
                if (prefData.sessionID && integration.sessions[prefData.sessionID]) {
                    prefData.sessionID = Zotero.Utilities.randomString(8);
                    await doc.setDocumentData(prefData.serialize());
                }
                if (!prefs(data)) return null;
                await recordTiming({ setDocPrefsRequired: false, styleConfirmedAt: new Date().toISOString(),
                    fieldsReadAt: current.readAt });
                await signal("settings", { needed: false, document: cfg.document, style: cfg.style });
                return refresh(doc, current);
            });
            if (!rendered) {
                await configureInWord();
                rendered = await attach("after-preferences", async (doc, current, data) => {
                    check(prefs(data), "Document Preferences closed without retaining the requested style and Word Fields");
                    if (!refreshTiming.styleConfirmedAt) {
                        await recordTiming({ styleConfirmedAt: new Date().toISOString(), styleConfirmationSource: "documentReadback" });
                    }
                    return refresh(doc, current);
                });
            }
            await timingWrite;
            await signal("refreshed", { ...rendered, origin: "refresh", scriptVersion: winWordScriptVersion, refreshTiming });
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
                const verified = await compareStyle(doc, saved);
                check(!verified.difference, verified.difference);
                entryCount = verified.entryCount;
            }
            for (const [index, id] of cfg.samples.entries()) {
                await handshake("sample-" + index, { citation: id, bibliographyCount: 1 });
                await confirmSelectedField(doc, id);
                stage = "editCitation " + id;
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
        await signal("verified", { executor: winWordExecutorVersion, document: cfg.document, sha256: cfg.documentHash, citations: expected.size,
            bibliographyEntries: entryCount, samples, style: cfg.style });
        return true;
    } catch (error) {
        if (refreshTiming) await timingWrite;
        await write("error.json", { stage, type: error.name, error: String(error), refreshTiming });
        return false;
    }
}
