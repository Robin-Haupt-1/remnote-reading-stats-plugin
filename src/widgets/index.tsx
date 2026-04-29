import {
  BuiltInPowerupCodes,
  declareIndexPlugin,
  type RNPlugin,
  type ReactRNPlugin,
  type Rem,
} from '@remnote/plugin-sdk';

type FocusSnapshot = {
  observedAt: string;
  documentHasFocus: boolean;
  visibilityState: DocumentVisibilityState;
  hidden: boolean;
  windowFocused: boolean | 'unknown';
  lastActivityAt: number | null;
  lastActivityType: string | null;
  millisSinceLastActivity: number | null;
};

type OpenSurfaceSnapshot = {
  source: 'window.getFocusedPaneId + window.getOpenPaneRemId';
  paneId?: string;
  remId?: string;
  remText?: string;
  kind: 'document' | 'pdf' | 'file' | 'rem' | 'none' | 'unknown';
  isDocument?: boolean;
  hasUploadedFile?: boolean;
  hasPdfHighlight?: boolean;
  fileName?: string;
  fileTitle?: string;
  fileType?: string;
  fileUrl?: string;
};

let lastKnownWindowFocus: boolean | 'unknown' = 'unknown';
let lastActivityAt: number | null = null;
let lastActivityType: string | null = null;
let cleanupListeners: (() => void) | undefined;
let stopOpenDocumentWatcher: (() => void) | undefined;

let lastOpenDocumentKey: string | null = null;

function logDetective(title: string, payload?: unknown): void {
  const prefix = '🕵️ [RemNote Detective]';
  if (payload === undefined) {
    console.log(`${prefix} ${title}`);
  } else {
    console.log(`${prefix} ${title}`, payload);
  }
}

async function toast(plugin: RNPlugin, message: string): Promise<void> {
  try {
    console.log('creating toast');
    await plugin.app.toast(message);
  } catch (error) {
    logDetective('Toast failed', error);
  }
}

function shortText(value: string | undefined, max = 55): string {
  if (!value) return 'n/a';
  return value.length <= max ? value : `${value.slice(0, max - 3)}...`;
}

function markActivity(type: string): void {
  lastActivityAt = Date.now();
  lastActivityType = type;
}

async function stringifyRemText(plugin: RNPlugin, rem: Rem | undefined | null): Promise<string | undefined> {
  if (!rem) return undefined;
  try {
    return await plugin.richText.toString(rem.text);
  } catch {
    return undefined;
  }
}

async function detectFocusSnapshot(_: RNPlugin): Promise<FocusSnapshot> {
  const now = Date.now();

  return {
    observedAt: new Date(now).toISOString(),
    documentHasFocus: document.hasFocus(),
    visibilityState: document.visibilityState,
    hidden: document.hidden,
    windowFocused: lastKnownWindowFocus,
    lastActivityAt,
    lastActivityType,
    millisSinceLastActivity: lastActivityAt == null ? null : now - lastActivityAt,
  };
}

async function detectOpenSurface(plugin: RNPlugin): Promise<OpenSurfaceSnapshot> {
  const paneId = await plugin.window.getFocusedPaneId().catch(() => undefined);
  const remId = paneId
    ? await plugin.window.getOpenPaneRemId(paneId).catch(() => undefined)
    : undefined;

  if (!remId) {
    return {
      source: 'window.getFocusedPaneId + window.getOpenPaneRemId',
      paneId,
      remId,
      kind: 'none',
    };
  }

  try {
    const rem = await plugin.rem.findOne(remId);

    if (!rem) {
      return {
        source: 'window.getFocusedPaneId + window.getOpenPaneRemId',
        paneId,
        remId,
        kind: 'unknown',
      };
    }

    const remText = await stringifyRemText(plugin, rem);
    const isDocument = await rem.isDocument().catch(() => false);
    const hasUploadedFile = await rem.hasPowerup(BuiltInPowerupCodes.UploadedFile).catch(() => false);
    const hasPdfHighlight = await rem.hasPowerup(BuiltInPowerupCodes.PDFHighlight).catch(() => false);

    let fileName: string | undefined;
    let fileTitle: string | undefined;
    let fileType: string | undefined;
    let fileUrl: string | undefined;
    let kind: OpenSurfaceSnapshot['kind'] = 'rem';

    if (hasUploadedFile) {
      fileName = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Name').catch(() => undefined);
      fileTitle = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Title').catch(() => undefined);
      fileType = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Type').catch(() => undefined);
      fileUrl = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'URL').catch(() => undefined);
    }

    if (hasPdfHighlight) {
      kind = 'pdf';
    } else if (hasUploadedFile) {
      const looksLikePdf =
        /pdf/i.test(fileType ?? '') ||
        /\.pdf($|\?)/i.test(fileUrl ?? '') ||
        /pdf/i.test(fileName ?? '');
      kind = looksLikePdf ? 'pdf' : 'file';
    } else if (isDocument) {
      kind = 'document';
    }

    return {
      source: 'window.getFocusedPaneId + window.getOpenPaneRemId',
      paneId,
      remId,
      remText,
      kind,
      isDocument,
      hasUploadedFile,
      hasPdfHighlight,
      fileName,
      fileTitle,
      fileType,
      fileUrl,
    };
  } catch (error) {
    logDetective('detectOpenSurface failed', error);
    return {
      source: 'window.getFocusedPaneId + window.getOpenPaneRemId',
      paneId,
      remId,
      kind: 'unknown',
    };
  }
}

function computeAppHasFocus(focus: FocusSnapshot): boolean {
  return (
    focus.visibilityState === 'visible' &&
    !focus.hidden &&
    (focus.documentHasFocus || focus.windowFocused === true)
  );
}

function getOpenDocumentName(open: OpenSurfaceSnapshot): string {
  return open.fileName ?? open.fileTitle ?? open.remText ?? 'n/a';
}

function getOpenDocumentKey(open: OpenSurfaceSnapshot): string {
  return `${open.remId ?? 'none'}::${getOpenDocumentName(open)}::${open.kind}`;
}

async function maybeToastOpenDocumentChange(
  plugin: RNPlugin,
  open: OpenSurfaceSnapshot,
  reason: string,
): Promise<void> {
  const nextKey = getOpenDocumentKey(open);
  const activeName = getOpenDocumentName(open);

  console.log('open document', activeName);

  if (nextKey !== lastOpenDocumentKey) {
    const previousKey = lastOpenDocumentKey;
    lastOpenDocumentKey = nextKey;

    logDetective('OPEN DOCUMENT CHANGED', {
      reason,
      previousKey,
      nextKey,
      name: activeName,
      remId: open.remId,
      kind: open.kind,
      paneId: open.paneId,
    });

    await toast(
      plugin,
      `🕵️ open changed | file=${shortText(activeName)} | kind=${open.kind} | rem=${open.remId ?? 'n/a'}`,
    );
  }
}

async function checkOpenDocumentChange(plugin: RNPlugin, reason: string): Promise<void> {
  const open = await detectOpenSurface(plugin);
  await maybeToastOpenDocumentChange(plugin, open, reason);
}

async function toastDefaultDebug(plugin: RNPlugin, reason: string): Promise<void> {
  const [focus, open] = await Promise.all([
    detectFocusSnapshot(plugin),
    detectOpenSurface(plugin),
  ]);

  const activeName = getOpenDocumentName(open);
  const appHasFocus = computeAppHasFocus(focus);

  logDetective(`DEBUG ${reason} | focus`, focus);
  logDetective(`DEBUG ${reason} | open`, open);

  await maybeToastOpenDocumentChange(plugin, open, reason);

  await toast(
    plugin,
    `🕵️ ${reason} | file=${shortText(activeName)} | appFocus=${String(appHasFocus)} | vis=${focus.visibilityState} | hasFocus=${String(
      focus.documentHasFocus,
    )}`,
  );
}

function installActivityAndFocusListeners(plugin: RNPlugin): () => void {
  const onWindowFocus: EventListener = () => {
    lastKnownWindowFocus = true;
    void toastDefaultDebug(plugin, 'window focus');
  };

  const onWindowBlur: EventListener = () => {
    lastKnownWindowFocus = false;
    void toastDefaultDebug(plugin, 'window blur');
  };

  const onVisibilityChange: EventListener = () => {
    void toastDefaultDebug(plugin, `visibility:${document.visibilityState}`);
  };

  const onPageShow: EventListener = () => {
    void toastDefaultDebug(plugin, 'pageshow');
  };

  const onPageHide: EventListener = () => {
    void toastDefaultDebug(plugin, 'pagehide');
  };

  const onScroll: EventListener = () => {
    markActivity('scroll');
  };

  const onMouseMove: EventListener = () => {
    markActivity('mousemove');
  };

  const onMouseDown: EventListener = () => {
    markActivity('mousedown');
  };

  const onKeyDown: EventListener = () => {
    markActivity('keydown');
  };

  const onTouchStart: EventListener = () => {
    markActivity('touchstart');
  };

  const onTouchMove: EventListener = () => {
    markActivity('touchmove');
  };

  const onPointerDown: EventListener = () => {
    markActivity('pointerdown');
  };

  const onPointerMove: EventListener = () => {
    markActivity('pointermove');
  };

  window.addEventListener('focus', onWindowFocus, true);
  window.addEventListener('blur', onWindowBlur, true);
  window.addEventListener('pageshow', onPageShow, true);
  window.addEventListener('pagehide', onPageHide, true);
  document.addEventListener('visibilitychange', onVisibilityChange, true);

  document.addEventListener('scroll', onScroll, { capture: true, passive: true });
  document.addEventListener('mousemove', onMouseMove, { capture: true, passive: true });
  document.addEventListener('mousedown', onMouseDown, { capture: true, passive: true });
  document.addEventListener('keydown', onKeyDown, true);
  document.addEventListener('touchstart', onTouchStart, { capture: true, passive: true });
  document.addEventListener('touchmove', onTouchMove, { capture: true, passive: true });
  document.addEventListener('pointerdown', onPointerDown, { capture: true, passive: true });
  document.addEventListener('pointermove', onPointerMove, { capture: true, passive: true });

  return () => {
    window.removeEventListener('focus', onWindowFocus, true);
    window.removeEventListener('blur', onWindowBlur, true);
    window.removeEventListener('pageshow', onPageShow, true);
    window.removeEventListener('pagehide', onPageHide, true);
    document.removeEventListener('visibilitychange', onVisibilityChange, true);

    document.removeEventListener('scroll', onScroll, { capture: true });
    document.removeEventListener('mousemove', onMouseMove, { capture: true });
    document.removeEventListener('mousedown', onMouseDown, { capture: true });
    document.removeEventListener('keydown', onKeyDown, true);
    document.removeEventListener('touchstart', onTouchStart, { capture: true });
    document.removeEventListener('touchmove', onTouchMove, { capture: true });
    document.removeEventListener('pointerdown', onPointerDown, { capture: true });
    document.removeEventListener('pointermove', onPointerMove, { capture: true });
  };
}

function startOpenDocumentWatcher(plugin: RNPlugin, intervalMs = 750): () => void {
  let stopped = false;
  let running = false;

  const tick = async () => {
    if (stopped || running) return;
    running = true;

    try {
      await checkOpenDocumentChange(plugin, 'interval watcher');
    } catch (error) {
      logDetective('Open document watcher tick failed', error);
    } finally {
      running = false;
    }
  };

  const intervalId = window.setInterval(() => {
    void tick();
  }, intervalMs);

  // immediate first check
  void tick();

  return () => {
    stopped = true;
    window.clearInterval(intervalId);
  };
}

async function onActivate(plugin: ReactRNPlugin) {
  const pluginId = plugin.manifest?.id ?? 'remnote-focus-detective';

  lastKnownWindowFocus = document.hasFocus();
  markActivity('activate');

  cleanupListeners = installActivityAndFocusListeners(plugin);
  stopOpenDocumentWatcher = startOpenDocumentWatcher(plugin, 750);

  await plugin.app.registerCommand({
    id: `${pluginId}:probe-now`,
    name: 'Run Probe Now',
    action: async () => {
      await toastDefaultDebug(plugin, 'manual probe');
    },
  });

  await plugin.app.registerCommand({
    id: `${pluginId}:show-focus-now`,
    name: 'Show Focus Now',
    action: async () => {
      const focus = await detectFocusSnapshot(plugin);
      const appHasFocus = computeAppHasFocus(focus);

      logDetective('FOCUS NOW', focus);
      await toast(
        plugin,
        `🕵️ focus=${String(appHasFocus)} | vis=${focus.visibilityState} | hasFocus=${String(
          focus.documentHasFocus,
        )} | last=${focus.lastActivityType ?? 'n/a'} | idleMs=${String(focus.millisSinceLastActivity)}`,
      );
    },
  });

  plugin.track(async (reactivePlugin) => {
    await reactivePlugin.window.getFocusedPaneId().catch(() => undefined);
    await reactivePlugin.window.getOpenPaneRemIds().catch(() => []);

    await checkOpenDocumentChange(reactivePlugin, 'reactive pane change');
    await toastDefaultDebug(reactivePlugin, 'reactive pane change');
  });

  await checkOpenDocumentChange(plugin, 'activation');
  await toastDefaultDebug(plugin, 'activation');
}

async function onDeactivate(_: ReactRNPlugin) {
  cleanupListeners?.();
  cleanupListeners = undefined;

  stopOpenDocumentWatcher?.();
  stopOpenDocumentWatcher = undefined;
}

declareIndexPlugin(onActivate, onDeactivate);