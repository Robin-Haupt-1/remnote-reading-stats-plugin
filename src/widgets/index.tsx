import {
  BuiltInPowerupCodes,
  declareIndexPlugin,
  type ReactRNPlugin,
  type Rem,
  type RNPlugin,
} from '@remnote/plugin-sdk';

let cleanupListeners: (() => void) | undefined;
let stopOpenDocumentWatcher: (() => void) | undefined;
let lastDocumentKey: string | null = null;

function logMessage(message: string): void {
  console.log(`[Reading Stats] ${message}`);
}

function installVisibilityListeners(): () => void {
  const onVisibilityRelatedEvent: EventListener = () => {
    logMessage(`visibility changed: ${document.visibilityState}`);
  };

  window.addEventListener('focus', onVisibilityRelatedEvent, true);
  window.addEventListener('blur', onVisibilityRelatedEvent, true);
  window.addEventListener('pageshow', onVisibilityRelatedEvent, true);
  window.addEventListener('pagehide', onVisibilityRelatedEvent, true);
  document.addEventListener('visibilitychange', onVisibilityRelatedEvent, true);

  return () => {
    window.removeEventListener('focus', onVisibilityRelatedEvent, true);
    window.removeEventListener('blur', onVisibilityRelatedEvent, true);
    window.removeEventListener('pageshow', onVisibilityRelatedEvent, true);
    window.removeEventListener('pagehide', onVisibilityRelatedEvent, true);
    document.removeEventListener('visibilitychange', onVisibilityRelatedEvent, true);
  };
}

async function getOpenUploadedFileRem(plugin: RNPlugin): Promise<Rem | null> {
  const paneId = await plugin.window.getFocusedPaneId().catch(() => undefined);
  const remId = paneId
    ? await plugin.window.getOpenPaneRemId(paneId).catch(() => undefined)
    : undefined;

  if (!remId) return null;

  const rem = await plugin.rem.findOne(remId);
  if (!rem) return null;

  const hasUploadedFile = await rem.hasPowerup(BuiltInPowerupCodes.UploadedFile).catch(() => false);
  return hasUploadedFile ? rem : null;
}

async function getOpenUploadedFileName(plugin: RNPlugin): Promise<string | null> {
  const rem = await getOpenUploadedFileRem(plugin);
  if (!rem) return null;

  const name = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Name').catch(() => undefined);
  return typeof name === 'string' && name.length > 0 ? name : 'unknown';
}

async function checkOpenDocumentChange(plugin: RNPlugin): Promise<void> {
  const name = await getOpenUploadedFileName(plugin);
  const nextKey = name ?? 'none';

  if (nextKey === lastDocumentKey) return;

  lastDocumentKey = nextKey;

  if (name) {
    logMessage(`open document changed: ${name}`);
  } else {
    logMessage('no document open');
  }
}

function startOpenDocumentWatcher(plugin: RNPlugin, intervalMs = 750): () => void {
  let stopped = false;
  let running = false;

  const tick = async () => {
    if (stopped || running) return;
    running = true;
    try {
      await checkOpenDocumentChange(plugin);
    } finally {
      running = false;
    }
  };

  const intervalId = window.setInterval(() => {
    void tick();
  }, intervalMs);

  void tick();

  return () => {
    stopped = true;
    window.clearInterval(intervalId);
  };
}

async function onActivate(plugin: ReactRNPlugin) {
  cleanupListeners = installVisibilityListeners();
  stopOpenDocumentWatcher = startOpenDocumentWatcher(plugin, 750);
}

async function onDeactivate(_: ReactRNPlugin) {
  cleanupListeners?.();
  cleanupListeners = undefined;

  stopOpenDocumentWatcher?.();
  stopOpenDocumentWatcher = undefined;

  lastDocumentKey = null;
}

declareIndexPlugin(onActivate, onDeactivate);
