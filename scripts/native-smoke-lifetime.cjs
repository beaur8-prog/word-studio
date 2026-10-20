'use strict';
// These applications use disposable test profiles, never the user's session.
// A native unsaved-changes prompt during cleanup must not hide an assertion
// failure or keep an unattended CI runner alive indefinitely.
module.exports = async function nativeSmokeLifetime(application, label) {
  const child = application.process();
  const terminate = () => { if (child.exitCode === null) child.kill('SIGKILL'); };
  const watchdog = setTimeout(() => {
    console.error(`${label}: isolated native test exceeded four minutes`);
    terminate();
    process.exit(1);
  }, 240000);
  watchdog.unref();
  await application.evaluate(({ dialog }) => {
    const original = dialog.showMessageBoxSync.bind(dialog);
    dialog.showMessageBoxSync = (...args) => {
      const options = args.at(-1);
      if (options?.title === 'Close Word Studio?') return 1;
      return original(...args);
    };
  });
  return async () => {
    let timeout;
    try {
      await Promise.race([
        application.close(),
        new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error(`${label}: test application did not close`)), 15000); }),
      ]);
    } catch (error) { terminate(); throw error; }
    finally { clearTimeout(timeout); clearTimeout(watchdog); }
  };
};
