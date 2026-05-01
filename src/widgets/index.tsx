import {
  BuiltInPowerupCodes,
  declareIndexPlugin,
  type ReactRNPlugin,
  type Rem,
} from '@remnote/plugin-sdk';
import '../style.css';
import '../index.css';

type HighlightRow = {
  fileName: string;
  fullText: string;
  pageNumber: string;
  updatedTimestamp: string;
};

type SlotProbeResult = {
  remId: string;
  remText: string;
  slot: string;
  ok: boolean;
  value: string;
};

const LOG_PREFIX = '[UploadedFile Slot Demo]';

const UPLOADED_FILE_SLOTS_TO_PROBE = [
  // Known/observed working slots.
  'Name',
  'Title',
  'URL',
  'Type',
  'Authors',
  'ViewerData',

  // Observed problematic slots.
  'ReadPercent',
  'LastReadDate',
  'Theme',
  'HasNoTextLayer',
];

function logMessage(message: string): void {
  console.log(`${LOG_PREFIX} ${message}`);
}

function safeJson(value: unknown): string {
  try {
    const json = JSON.stringify(value);
    if (!json) return String(value);
    return json.length > 700 ? `${json.slice(0, 700)}… <truncated>` : json;
  } catch {
    return String(value);
  }
}

async function remTextToString(plugin: ReactRNPlugin, rem: Rem): Promise<string> {
  try {
    return await plugin.richText.toString(rem.text);
  } catch {
    return '<text unavailable>';
  }
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildExcelXml(rows: HighlightRow[]): string {
  const header = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <Worksheet ss:Name="Highlights">
    <Table>
      <Row>
        <Cell><Data ss:Type="String">File Name</Data></Cell>
        <Cell><Data ss:Type="String">Full Text</Data></Cell>
        <Cell><Data ss:Type="String">Page Number</Data></Cell>
        <Cell><Data ss:Type="String">Updated Timestamp</Data></Cell>
      </Row>`;

  const body = rows
    .map(
      (row) => `
      <Row>
        <Cell><Data ss:Type="String">${escapeXml(row.fileName)}</Data></Cell>
        <Cell><Data ss:Type="String">${escapeXml(row.fullText)}</Data></Cell>
        <Cell><Data ss:Type="String">${escapeXml(row.pageNumber)}</Data></Cell>
        <Cell><Data ss:Type="String">${escapeXml(row.updatedTimestamp)}</Data></Cell>
      </Row>`,
    )
    .join('');

  const footer = `
    </Table>
  </Worksheet>
</Workbook>`;

  return `${header}${body}${footer}`;
}

function downloadExcel(xml: string): void {
  const encoded = btoa(unescape(encodeURIComponent(xml)));
  const href = `data:application/vnd.ms-excel;base64,${encoded}`;
  const now = new Date().toISOString().replace(/[:.]/g, '-');

  const link = document.createElement('a');
  link.href = href;
  link.download = `remnote-highlights-${now}.xls`;
  link.target = '_blank';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

async function collectHighlightsForRem(
  rem: Rem,
  plugin: ReactRNPlugin,
  fileName: string,
  rows: HighlightRow[],
): Promise<void> {
  const children = await rem.getChildrenRem();

  for (const child of children) {
    const isPdfHighlight = await child.hasPowerup(BuiltInPowerupCodes.PDFHighlight);

    if (isPdfHighlight) {
      const fullText = await plugin.richText.toString(child.text);
      const pageNumber = await plugin.richText.toString((await child.getParentRem())?.text ?? []);
      const updatedTimestamp = new Date(child.updatedAt).toISOString();

      rows.push({
        fileName,
        fullText,
        pageNumber,
        updatedTimestamp,
      });
    }

    await collectHighlightsForRem(child, plugin, fileName, rows);
  }
}

async function getUploadedFileRems(plugin: ReactRNPlugin): Promise<Rem[]> {
  const filePowerup = await plugin.powerup.getPowerupByCode(BuiltInPowerupCodes.UploadedFile);

  if (!filePowerup) {
    return [];
  }

  const taggedRems = await filePowerup.taggedRem();
  const uploadedFileRems: Rem[] = [];

  for (const rem of taggedRems) {
    const confirmed = await rem.hasPowerup(BuiltInPowerupCodes.UploadedFile).catch(() => false);

    if (confirmed) {
      uploadedFileRems.push(rem);
    }
  }

  return uploadedFileRems;
}

async function getUploadedFileName(plugin: ReactRNPlugin, rem: Rem): Promise<string> {
  const name = await rem
    .getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Name')
    .catch(() => undefined);

  if (typeof name === 'string' && name.length > 0) {
    return name;
  }

  return (await remTextToString(plugin, rem)) || 'Unknown File';
}

async function exportHighlights(plugin: ReactRNPlugin): Promise<void> {
  await plugin.app.toast('Exporting PDF highlights to Excel...');

  const pdfRems = await getUploadedFileRems(plugin);

  if (pdfRems.length === 0) {
    await plugin.app.toast('No uploaded file Rems found.');
    return;
  }

  const rows: HighlightRow[] = [];

  for (const pdfRem of pdfRems) {
    const fileName = await getUploadedFileName(plugin, pdfRem);
    await collectHighlightsForRem(pdfRem, plugin, fileName, rows);
  }

  if (rows.length === 0) {
    await plugin.app.toast('No PDF highlights found to export.');
    return;
  }

  const xml = buildExcelXml(rows);
  downloadExcel(xml);

  await plugin.app.toast(`Excel download started for ${rows.length} highlights.`);
}

async function probeUploadedFileSlot(
  plugin: ReactRNPlugin,
  rem: Rem,
  slot: string,
): Promise<SlotProbeResult> {
  const remText = await remTextToString(plugin, rem);

  try {
    const value = await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, slot);

    return {
      remId: rem._id,
      remText,
      slot,
      ok: true,
      value: safeJson(value),
    };
  } catch (error) {
    return {
      remId: rem._id,
      remText,
      slot,
      ok: false,
      value: error instanceof Error ? error.message : String(error),
    };
  }
}

async function demoUploadedFileSlotReads(plugin: ReactRNPlugin): Promise<void> {
  await plugin.app.toast('Running UploadedFile slot demo. Check the console.');

  const uploadedFileRems = await getUploadedFileRems(plugin);

  if (uploadedFileRems.length === 0) {
    logMessage('No Rems found that are confirmed UploadedFile Rems.');
    await plugin.app.toast('No confirmed UploadedFile Rems found.');
    return;
  }

  logMessage(`Found ${uploadedFileRems.length} confirmed UploadedFile Rem(s).`);
  logMessage(`BuiltInPowerupCodes.UploadedFile = ${safeJson(BuiltInPowerupCodes.UploadedFile)}`);

  const maxRemsToProbe = 5;
  const remsToProbe = uploadedFileRems.slice(0, maxRemsToProbe);

  let successCount = 0;
  let failureCount = 0;

  for (const rem of remsToProbe) {
    const remText = await remTextToString(plugin, rem);
    const confirmed = await rem.hasPowerup(BuiltInPowerupCodes.UploadedFile).catch(() => false);

    logMessage('------------------------------------------------------------');
    logMessage(`Rem id=${rem._id}`);
    logMessage(`Confirmed UploadedFile=${confirmed}`);
    logMessage(`Rem text=${remText}`);

    if (!confirmed) {
      logMessage('Skipping because this Rem is not confirmed as UploadedFile.');
      continue;
    }

    for (const slot of UPLOADED_FILE_SLOTS_TO_PROBE) {
      const result = await probeUploadedFileSlot(plugin, rem, slot);

      if (result.ok) {
        successCount += 1;
        logMessage(
          `OK slot="${result.slot}" value=${result.value}`,
        );
      } else {
        failureCount += 1;
        logMessage(
          `FAIL slot="${result.slot}" error=${safeJson(result.value)}`,
        );
      }
    }
  }

  logMessage('------------------------------------------------------------');
  logMessage(
    `Demo complete. Probed ${remsToProbe.length}/${uploadedFileRems.length} UploadedFile Rem(s). Successes=${successCount}, Failures=${failureCount}`,
  );

  await plugin.app.toast(
    `UploadedFile slot demo complete. Successes=${successCount}, Failures=${failureCount}. Check console.`,
  );
}

async function onActivate(plugin: ReactRNPlugin) {
  const pluginId = plugin.manifest?.id ?? 'remnote-export-highlights-plugin';

  await plugin.app.registerCommand({
    id: `${pluginId}:export-highlights-excel`,
    name: 'Export PDF Highlights to Excel',
    description: 'Export all PDF highlights with file name, text, page number, and updated timestamp.',
    keywords: 'pdf highlight export excel xls',
    action: async () => {
      await exportHighlights(plugin);
    },
  });

  await plugin.app.registerCommand({
    id: `${pluginId}:demo-uploaded-file-slot-reads`,
    name: 'Demo UploadedFile Slot Reads',
    description:
      'Replicate UploadedFile powerup slot reads and log working/misleading failing slots to the console.',
    keywords: 'uploaded file powerup slot readpercent viewerdata debug demo',
    action: async () => {
      await demoUploadedFileSlotReads(plugin);
    },
  });
}

async function onDeactivate(_: ReactRNPlugin) {}

declareIndexPlugin(onActivate, onDeactivate);