import {
  BuiltInPowerupCodes,
  declareIndexPlugin,
  type ReactRNPlugin,
  type Rem,
  type RNPlugin,
} from '@remnote/plugin-sdk';

type EndReason = 'visibility' | 'closed' | 'switched' | 'deactivate';

type ReadingEvent = {
  startedTimestamp: number;
  finishedTimestamp: number;
  documentName: string;
  endedReason: EndReason;
};

type ReadingSessionState = {
  current: {
    startedTimestamp: number;
    documentName: string;
  } | null;
  events: ReadingEvent[];
};

const READING_STATE_KEY = 'reading-stats:events-v1';
const TICK_MS = 1000;

let cleanupListeners: (() => void) | undefined;
let stopOpenDocumentWatcher: (() => void) | undefined;
let stopStateFlushTicker: (() => void) | undefined;
let sessionState: ReadingSessionState = { current: null, events: [] };
let stateDirty = false;
let flushingState = false;

function logMessage(message: string): void {
  console.log(`[Reading Stats] ${message}`);
}

function parseReadingState(value: unknown): ReadingSessionState {
  if (!value || typeof value !== 'object') return { current: null, events: [] };

  const maybeState = value as Partial<ReadingSessionState>;
  const events = Array.isArray(maybeState.events)
    ? maybeState.events.filter(
        (event): event is ReadingEvent =>
          Boolean(
            event &&
              typeof event.startedTimestamp === 'number' &&
              typeof event.finishedTimestamp === 'number' &&
              typeof event.documentName === 'string' &&
              typeof event.endedReason === 'string',
          ),
      )
    : [];

  const current =
    maybeState.current &&
    typeof maybeState.current.startedTimestamp === 'number' &&
    typeof maybeState.current.documentName === 'string'
      ? {
          startedTimestamp: maybeState.current.startedTimestamp,
          documentName: maybeState.current.documentName,
        }
      : null;

  return { current, events };
}

async function loadReadingState(plugin: RNPlugin): Promise<void> {
  const value = await plugin.storage.getSynced(READING_STATE_KEY).catch(() => undefined);
  sessionState = parseReadingState(value);
}

async function flushReadingState(plugin: RNPlugin, force = false): Promise<void> {
  if ((!stateDirty && !force) || flushingState) return;

  flushingState = true;
  try {
    await plugin.storage.setSynced(READING_STATE_KEY, sessionState);
    stateDirty = false;
  } finally {
    flushingState = false;
  }
}

function markDirty(): void {
  stateDirty = true;
}

function closeCurrentEvent(reason: EndReason): void {
  if (!sessionState.current) return;

  sessionState.events.push({
    startedTimestamp: sessionState.current.startedTimestamp,
    finishedTimestamp: Date.now(),
    documentName: sessionState.current.documentName,
    endedReason: reason,
  });
  sessionState.current = null;
  markDirty();
}

function startCurrentEvent(documentName: string): void {
  sessionState.current = {
    startedTimestamp: Date.now(),
    documentName,
  };
  markDirty();
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

  if (!sessionState.current) {
    if (name) {
      startCurrentEvent(name);
      logMessage(`started reading: ${name}`);
    }
    return;
  }

  if (!name) {
    closeCurrentEvent('closed');
    logMessage('reading ended: document closed');
    return;
  }

  if (sessionState.current.documentName !== name) {
    closeCurrentEvent('switched');
    startCurrentEvent(name);
    logMessage(`switched reading document: ${name}`);
  }
}

function installVisibilityListeners(): () => void {
  const onVisibilityChange = () => {
    if (document.visibilityState !== 'visible') {
      closeCurrentEvent('visibility');
      logMessage(`reading ended due to visibility change: ${document.visibilityState}`);
    }
  };

  window.addEventListener('blur', onVisibilityChange, true);
  window.addEventListener('pagehide', onVisibilityChange, true);
  document.addEventListener('visibilitychange', onVisibilityChange, true);

  return () => {
    window.removeEventListener('blur', onVisibilityChange, true);
    window.removeEventListener('pagehide', onVisibilityChange, true);
    document.removeEventListener('visibilitychange', onVisibilityChange, true);
  };
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

function startStateFlushTicker(plugin: RNPlugin, intervalMs = TICK_MS): () => void {
  const intervalId = window.setInterval(() => {
    void flushReadingState(plugin);
  }, intervalMs);

  return () => {
    window.clearInterval(intervalId);
  };
}

function csvEscape(value: string): string {
  const escaped = value.replace(/"/g, '""');
  return /[",\n]/.test(value) ? `"${escaped}"` : escaped;
}

function downloadCsv(csv: string): void {
  const now = new Date().toISOString().replace(/[:.]/g, '-');
  const href = `data:text/csv;charset=utf-8,${encodeURIComponent(csv)}`;

  const link = document.createElement('a');
  link.href = href;
  link.download = `remnote-reading-events-${now}.csv`;
  link.target = '_blank';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

async function exportReadingEvents(plugin: RNPlugin): Promise<void> {
  await flushReadingState(plugin, true);

  const persistedState = parseReadingState(await plugin.storage.getSynced(READING_STATE_KEY).catch(() => undefined));

  if (persistedState.events.length === 0) {
    await plugin.app.toast('No reading events found to export.');
    return;
  }

  const header = ['startedTimestamp', 'finishedTimestamp', 'documentName', 'endedReason'];
  const rows = persistedState.events.map((event) =>
    [
      new Date(event.startedTimestamp).toISOString(),
      new Date(event.finishedTimestamp).toISOString(),
      event.documentName,
      event.endedReason,
    ]
      .map(csvEscape)
      .join(','),
  );

  const csv = [header.join(','), ...rows].join('\n');
  downloadCsv(csv);
  await plugin.app.toast(`CSV download started for ${persistedState.events.length} reading events.`);
}

async function onActivate(plugin: ReactRNPlugin) {
  await loadReadingState(plugin);

  cleanupListeners = installVisibilityListeners();
  stopOpenDocumentWatcher = startOpenDocumentWatcher(plugin, 750);
  stopStateFlushTicker = startStateFlushTicker(plugin, TICK_MS);

  const pluginId = plugin.manifest?.id ?? 'remnote-reading-stats-plugin';
  await plugin.app.registerCommand({
    id: `${pluginId}:export-reading-events-csv`,
    name: 'Export Reading Events to CSV',
    description: 'Export tracked reading events (start, finish, document, ended reason) to CSV.',
    keywords: 'reading stats export csv events',
    action: async () => {
      await exportReadingEvents(plugin);
    },
  });
}

async function onDeactivate(plugin: ReactRNPlugin) {
  cleanupListeners?.();
  cleanupListeners = undefined;

  stopOpenDocumentWatcher?.();
  stopOpenDocumentWatcher = undefined;

  stopStateFlushTicker?.();
  stopStateFlushTicker = undefined;

  closeCurrentEvent('deactivate');
  await flushReadingState(plugin, true);
}

declareIndexPlugin(onActivate, onDeactivate);
