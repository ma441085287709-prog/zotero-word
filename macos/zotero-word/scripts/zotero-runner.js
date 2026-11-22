/* macOS Accessibility opens Zotero's own Run JavaScript menu. The only text
 * entered is the loader for the bundled script; job data stays in its JSON file. */
ObjC.import("AppKit");
ObjC.import("CoreGraphics");

const runLabels = ["Run JavaScript", "执行 JavaScript", "运行 JavaScript", "執行 JavaScript"];

function until(label, seconds, get) {
    const deadline = Date.now() + seconds * 1000;
    while (Date.now() < deadline) {
        try {
            const value = get();
            if (value) return value;
        } catch (error) {
            if (error.errorNumber !== -1728) throw error; // The AX object disappeared while the tree was being built.
        }
        delay(0.08);
    }
    throw Error("Timed out waiting for " + label);
}

function attribute(element, name) {
    const entry = element.attributes.byName(name);
    return entry.exists() ? entry.value() : null;
}

function identifiedAs(element, id) {
    return attribute(element, "AXIdentifier") === id || attribute(element, "AXDOMIdentifier") === id;
}

function findRunMenu(proc, observation) {
    const preferred = ["Tools", "工具", "Outils", "Werkzeuge", "Herramientas", "Strumenti", "ツール"];
    observation.menus = [];
    observation.tools = [];
    observation.developer = [];
    if (!proc.menuBars[0].exists()) return null;
    const tops = proc.menuBars[0].menuBarItems;
    const names = observation.menus = tops.name();
    if (names.length <= 2) return null;
    function command(menu, isDeveloper) {
        const labels = menu.menuItems.name();
        if (isDeveloper) observation.developer = labels;
        const known = labels.find(t => runLabels.includes(t));
        if (known) return menu.menuItems.byName(known);
        for (const label of labels.filter(t => /java\s?script/i.test(t || ""))) {
            const item = menu.menuItems.byName(label);
            if (identifiedAs(item, "menu_runJS")) return item;
        }
        return null;
    }
    const toolsMenus = names.filter(name => preferred.includes(name));
    for (const name of [...toolsMenus, ...names.filter(title => !preferred.includes(title))]) {
        if (!name) continue;
        const menu = tops.byName(name).menus[0];
        if (!menu.exists()) continue;
        const labels = menu.menuItems.name();
        if (toolsMenus.includes(name)) observation.tools = labels;
        const developer = labels.find(label => ["Developer", "开发者", "開發者"].includes(label));
        if (developer) {
            const submenu = menu.menuItems.byName(developer).menus[0];
            if (submenu.exists()) {
                const run = command(submenu, true);
                if (run) return run;
            }
        }
        for (const child of menu.menuItems()) {
            const submenu = child.menus[0];
            if (!submenu.exists()) continue;
            const run = command(submenu);
            if (run) return run;
        }
    }
    return null;
}

// Quartz supplies stable window numbers even when two windows have the same title.
function zoteroWindows(pid, option) {
    const list = ObjC.deepUnwrap(ObjC.castRefToObject($.CGWindowListCopyWindowInfo(option, 0)));
    if (!Array.isArray(list)) throw Error("Cannot read windows in this macOS desktop session.");
    const windows = [];
    for (const info of list) {
        if (info.kCGWindowOwnerPID === pid) {
            windows.push({id: info.kCGWindowNumber, layer: info.kCGWindowLayer, bounds: info.kCGWindowBounds});
        }
    }
    return windows;
}

function clickInPane(pid, bounds, fraction) {
    if (!bounds || ![bounds.X, bounds.Y, bounds.Width, bounds.Height].every(Number.isFinite) ||
            bounds.Width <= 0 || bounds.Height <= 0) throw Error("Cannot locate the script window on screen.");
    const point = $.CGPointMake(bounds.X + fraction * bounds.Width, bounds.Y + 0.45 * bounds.Height);
    for (const kind of [$.kCGEventLeftMouseDown, $.kCGEventLeftMouseUp]) {
        const event = $.CGEventCreateMouseEvent(null, kind, point, $.kCGMouseButtonLeft);
        $.CGEventSetIntegerValueField(event, $.kCGMouseEventClickState, 1);
        $.CGEventPostToPid(pid, event);
    }
}

function clipboardCopy(board) {
    const saved = $.NSMutableArray.alloc.init;
    const items = board.pasteboardItems;
    const count = Number(items.count);
    for (let i = 0; i < count; i++) {
        const original = items.objectAtIndex(i);
        const copy = $.NSPasteboardItem.alloc.init;
        const types = original.types;
        for (let j = 0; j < Number(types.count); j++) {
            const type = types.objectAtIndex(j);
            const data = original.dataForType(type);
            if (!data || data.isNil() || !copy.setDataForType(data, type)) {
                throw Error("The current clipboard cannot be preserved.");
            }
        }
        saved.addObject(copy);
    }
    return saved;
}

function shortcut(pid, key) {
    // Address both events to Zotero, even if another app takes the foreground.
    for (const down of [true, false]) {
        const event = $.CGEventCreateKeyboardEvent(null, key, down);
        $.CGEventSetFlags(event, 1 << 20); // kCGEventFlagMaskCommand
        $.CGEventPostToPid(pid, event);
    }
}

function openAndDispatch(loader, directory, observation, operation) {
    const bundle = "org.zotero.zotero";
    Application(bundle).activate();
    const se = Application("System Events");
    const processes = se.applicationProcesses.whose({bundleIdentifier: bundle});
    const activationDeadline = Date.now() + 15000;
    const menuDeadline = activationDeadline + 5000; // Includes process registration and the initial activation.
    let canRecoverFocus = operation === "refresh";
    const [proc, pid] = until("Zotero to register the same PID with macOS and System Events", 15, () => {
        const apps = $.NSRunningApplication.runningApplicationsWithBundleIdentifier($(bundle));
        const candidates = processes();
        observation.appCount = Number(apps.count);
        observation.accessibilityCount = candidates.length;
        observation.appPid = observation.appCount === 1 ? Number(apps.firstObject.processIdentifier) : null;
        observation.accessibilityPid = candidates.length === 1 ? candidates[0].unixId() : null;
        if (observation.appCount > 1 || candidates.length > 1) throw Error("Multiple Zotero desktop processes are open.");
        const pid = observation.appPid;
        if (!pid || observation.accessibilityPid !== pid) return null;
        const proc = candidates[0];
        observation.frontmost = proc.frontmost();
        if (!observation.frontmost && !observation.windows) observation.windows = proc.windows.name();
        return (observation.frontmost || canRecoverFocus) ? [proc, pid] : null;
    });
    function isFront() {
        if (proc.unixId() !== pid) throw Error("The Zotero process changed before dispatch.");
        observation.frontmost = proc.frontmost();
        return observation.frontmost;
    }
    function regainFront(deadline, mayReopenLibrary = false) {
        if (!canRecoverFocus || observation.runSent) return false;
        canRecoverFocus = false;
        observation.windows = proc.windows.name();
        if (mayReopenLibrary && !observation.windows.length) {
            // A running Zotero with no library window needs a reopen event.
            observation.focusRecovery = "reopen";
            $.NSTask.launchedTaskWithLaunchPathArguments($("/usr/bin/open"), $(["-b", bundle]));
        } else {
            observation.focusRecovery = "activate";
            Application(bundle).activate();
        }
        return until("the same Zotero process to return to the foreground", (deadline - Date.now()) / 1000, isFront);
    }
    // Bring the library forward if an auxiliary window has only the application menu.
    let raisedLibrary = false;
    let awaitingActivation = !observation.frontmost;
    const [entry, beforeTitles, previous] = until("the Run JavaScript menu in Zotero's main library window", Math.max(1, (menuDeadline - Date.now()) / 1000), () => {
        observation.phase = "menu";
        if (!isFront()) {
            if (awaitingActivation && Date.now() < activationDeadline) return null;
            if (!regainFront(menuDeadline, true)) {
                observation.windows = proc.windows.name();
                throw Error("Zotero lost focus while locating the Run JavaScript menu.");
            }
        }
        awaitingActivation = false;
        const menu = findRunMenu(proc, observation);
        observation.windows = proc.windows.name();
        if (menu) {
            observation.phase = "window";
            const titles = new Set(observation.windows);
            if (titles.has(menu.name()) || runLabels.some(title => titles.has(title))) {
                throw Error("A Zotero script editor is already open; preserve its contents and close it before resuming.");
            }
            const ids = new Set(zoteroWindows(pid, $.kCGWindowListOptionAll).map(w => w.id));
            // If focus changed during the lookup, reacquire the menu and its window snapshot together.
            return isFront() ? [menu, titles, ids] : null;
        }
        if (raisedLibrary) return null;
        const libraries = observation.windows.filter(name => /(?:^Zotero$| [-–—] Zotero$)/.test(name || ""));
        if (libraries.length === 1) {
            const library = proc.windows.byName(libraries[0]);
            const main = library.attributes.byName("AXMain");
            if (main.exists() && main.settable()) main.value = true;
            const raise = library.actions.byName("AXRaise");
            if (raise.exists()) se.perform(raise);
            raisedLibrary = true;
        }
        return null;
    });
    function newWindowId() {
        const fresh = zoteroWindows(pid, $.kCGWindowListOptionOnScreenOnly).filter(w => !previous.has(w.id));
        const ordinary = fresh.filter(w => w.layer === 0);
        const candidates = ordinary.length ? ordinary : fresh;
        return candidates.length === 1 ? candidates[0].id : null;
    }
    se.click(entry);
    const windowDeadline = Date.now() + 5000;
    const [title, id] = until("a new Zotero Run JavaScript window (close an existing script editor before resuming)", 5, () => {
        if (!isFront() && !regainFront(windowDeadline)) return null;
        const front = proc.windows[0];
        const title = front.exists() && front.name();
        if (!title || beforeTitles.has(title) || !/java\s?script/i.test(title)) return null;
        const id = newWindowId();
        return id && [title, id];
    });
    observation.windowId = id;
    observation.windowTitle = title;
    function requireWindow() {
        do {
            const titles = observation.windows = proc.windows.name();
            const window = zoteroWindows(pid, $.kCGWindowListOptionOnScreenOnly).find(w => w.id === id);
            if (isFront()) {
                if (titles[0] === title && titles.lastIndexOf(title) === 0 && window) return window.bounds;
                break;
            }
        } while (regainFront(Date.now() + 5000));
        throw Error("Zotero's newly opened script window is no longer in front.");
    }
    observation.phase = "clipboard";
    const board = $.NSPasteboard.generalPasteboard;
    const initial = Number(board.changeCount);
    const saved = clipboardCopy(board);
    const plain = $.NSPasteboardTypeString;
    const marker = $("org.zotero-word.loader-owner");
    const item = $.NSPasteboardItem.alloc.init;
    if (!item.setStringForType($(loader), plain) || !item.setStringForType($(directory), marker)) {
        throw Error("Cannot prepare the Zotero loader on the clipboard.");
    }
    function stringFor(type) {
        const value = board.stringForType(type);
        return value && !value.isNil() ? ObjC.unwrap(value) : null;
    }
    const canonical = s => typeof s === "string" ? s.replace(/\r\n?/g, "\n") : null;
    const wanted = canonical(loader);
    let echoed = false;
    requireWindow();
    if (Number(board.changeCount) !== initial) throw Error("The clipboard changed before dispatch.");
    let owned = Number(board.clearContents);
    function stillStaged() {
        return Number(board.changeCount) === owned && stringFor(marker) === directory && canonical(stringFor(plain)) === wanted;
    }
    function editorKey(key) {
        requireWindow();
        if (!stillStaged()) {
            observation.phase = "clipboard";
            throw Error("The clipboard changed while addressing the Zotero editor.");
        }
        shortcut(pid, key);
    }
    function readCopiedLoader() {
        const deadline = Date.now() + 1000;
        do {
            const generation = Number(board.changeCount);
            if (generation !== owned && stringFor(marker) === null && canonical(stringFor(plain)) === wanted &&
                    Number(board.changeCount) === generation) {
                owned = generation;
                return true;
            }
            delay(0.05);
        } while (Date.now() < deadline);
        if (!stillStaged()) {
            observation.phase = "clipboard";
            throw Error("The clipboard changed without an exact copy of the loader from the editor.");
        }
        return false;
    }
    try {
        if (Number(board.changeCount) !== owned || !board.writeObjects($.NSArray.arrayWithObject(item))) {
            throw Error("Cannot stage the Zotero loader on the clipboard.");
        }
        owned = Number(board.changeCount);
        observation.phase = "focus";
        // runJS.js focuses Ace on load. Its iframe can be absent from the entire macOS AX tree.
        // The two-pane form has only one writable editor. A click is a fallback; the copied code is the proof.
        for (const [method, fraction] of [["autofocus", null], ["left-pane", 0.25], ["right-pane", 0.75]]) {
            observation.focusMethod = method;
            if (fraction !== null) {
                clickInPane(pid, requireWindow(), fraction);
                editorKey(0); // Command-A, replace any earlier paste in our newly created window.
            }
            editorKey(9);     // Command-V
            editorKey(0);     // Command-A
            editorKey(8);     // Command-C
            echoed = readCopiedLoader();
            if (echoed) break;
        }
        if (!echoed) throw Error("The bound Zotero script window did not return the pasted loader when copied.");
        observation.loaderReadBack = true;
        requireWindow();
        if (Number(board.changeCount) !== owned || canonical(stringFor(plain)) !== wanted) {
            observation.phase = "clipboard";
            throw Error("The clipboard changed after the editor readback.");
        }
        observation.phase = "ack";
        observation.runSent = true;
        shortcut(pid, 15);    // Command-R, handled by the Run JavaScript window.
        const fs = $.NSFileManager.defaultManager;
        until("Zotero to acknowledge the bundled script", 12, () =>
            ["started.json", "error.json", "session-error.json"].some(name =>
                fs.fileExistsAtPath($(directory + "/" + name))));
    } finally {
        if (Number(board.changeCount) === owned &&
                (stillStaged() || (echoed && canonical(stringFor(plain)) === wanted) || !Number(board.pasteboardItems.count))) {
            board.clearContents;
            if (Number(saved.count) && !board.writeObjects(saved)) throw Error("Could not restore the previous clipboard.");
        }
    }
    return pid;
}

function run(argv) {
    const observation = {phase: "process", runSent: false};
    try {
        return JSON.stringify({pid: openAndDispatch(argv[0], argv[1], observation, argv[2])});
    } catch (error) {
        observation.type = error.name || "Error";
        observation.detail = error.message || String(error);
        observation.appleEvent = error.errorNumber || null;
        return JSON.stringify({error: observation});
    }
}
