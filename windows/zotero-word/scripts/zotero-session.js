/* One invocation of the bundled connector. Each handoff names one sibling job. */
async function serve(firstDir, runJob) {
    const root = Zotero.File.pathToFile(firstDir).parent.path;
    const session = Zotero.File.pathToFile(firstDir).leafName;
    const jobName = /^native-[0-9a-f]{32}$/;
    const stopped = Zotero.File.pathToFile(firstDir + "\\session.stop");
    let job = session;
    try {
        if (!jobName.test(job)) throw new Error("Invalid initial native job");
        while (!stopped.exists()) {
            const directory = root + "\\" + job;
            if (!await runJob(directory, session, stopped)) return;
            const nextPath = directory + "\\next.json";
            const next = Zotero.File.pathToFile(nextPath);
            const abort = Zotero.File.pathToFile(directory + "\\abort");
            const idleUntil = Date.now() + 120000;
            while (!next.exists()) {
                if (stopped.exists() || abort.exists()) return;
                if (Date.now() >= idleUntil) {
                    // Compete atomically with the controller's exclusive rename. An empty
                    // slot means we retired before accepting another request.
                    try {
                        next.create(0, 0o600);
                        return;
                    } catch (error) {
                        if (!next.exists()) throw error;
                    }
                    break;
                }
                await Zotero.Promise.delay(200);
            }
            if (stopped.exists() || abort.exists()) return;
            const request = JSON.parse(await Zotero.File.getContentsAsync(nextPath));
            if (request.session !== session || typeof request.job !== "string" || !jobName.test(request.job) || request.job === job) {
                throw new Error("The handoff does not belong to this Zotero session");
            }
            job = request.job;
        }
    } catch (error) {
        await Zotero.File.putContentsAsync(firstDir + "\\session-error.json",
            JSON.stringify({ job, session, type: error.name, error: String(error) }));
    } finally {
        await Zotero.File.putContentsAsync(firstDir + "\\session-ended.json", JSON.stringify({job, session}));
    }
}
