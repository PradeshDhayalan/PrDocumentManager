import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import sharp from 'sharp';
import JSZip from 'jszip';
import { Entity, Snapshot, records, userIds } from '../api/model';
import { Store } from '../api/store';
const root = path.resolve(__dirname, '../..');
const folder = path.join(root, 'seed-files');
const fixtures: [string, string][] = [
  [
    'Master Services Agreement.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ],
  ['Q3 Proposal.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
  ['Pricing Model.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
  ['Signed NDA.pdf', 'application/pdf'],
  ['Insurance Certificate.pdf', 'application/pdf'],
  ['Site Photo.jpg', 'image/jpeg'],
  ['Logo.png', 'image/png'],
  ['Kickoff Recording.mp4', 'video/mp4'],
  ['Meeting Notes.txt', 'text/plain'],
  ['Requirements.md', 'text/markdown'],
  ['Export.csv', 'text/csv'],
  ['config.json', 'application/json'],
  ['Archive.zip', 'application/zip'],
  ['Network Diagram.vsdx', 'application/vnd.ms-visio.drawing'],
];
const seedGuid = (group: number, index: number) =>
  `${String(group).padStart(8, '0')}-0000-4000-8000-${String(index).padStart(12, '0')}`;
export async function generateSeed(seedRoot = folder, now = Date.now()): Promise<Snapshot> {
  fs.mkdirSync(seedRoot, { recursive: true });
  const write = (name: string, data: Buffer | Uint8Array | string) =>
    fs.writeFileSync(path.join(seedRoot, name), data);
  const word = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: 'Master Services Agreement', heading: 'Heading1' }),
          new Paragraph({
            children: [
              new TextRun(
                'Contoso Ltd — sample agreement for the DMS Grid harness. This is test data.',
              ),
            ],
          }),
        ],
      },
    ],
  });
  write(fixtures[0][0], await Packer.toBuffer(word));
  const pptx = new PptxGenJS();
  pptx.author = 'DMS Grid sample generator';
  pptx.subject = 'Local test data';
  for (const title of ['Q3 Proposal', 'Scope and delivery']) {
    const slide = pptx.addSlide();
    slide.addText(title, { x: 0.7, y: 0.5, w: 8, h: 1, fontSize: 30, color: '0078D4' });
    slide.addText('Contoso Ltd • Sample document for UI and transfer testing', {
      x: 0.7,
      y: 2,
      w: 8,
      h: 1,
      fontSize: 18,
    });
  }
  await pptx.writeFile({ fileName: path.join(seedRoot, fixtures[1][0]) });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DMS Grid sample generator';
  const sheet = workbook.addWorksheet('Pricing');
  sheet.addRows([
    ['Item', 'Quantity', 'Unit price', 'Total'],
    ['Implementation', 1, 12500, { formula: 'B2*C2', result: 12500 }],
    ['Support', 12, 450, { formula: 'B3*C3', result: 5400 }],
  ]);
  sheet.columns.forEach((column) => {
    column.width = 22;
  });
  sheet.getRow(1).font = { bold: true };
  await workbook.xlsx.writeFile(path.join(seedRoot, fixtures[2][0]));
  for (const title of ['Signed NDA', 'Insurance Certificate']) {
    const pdf = await PDFDocument.create(),
      font = await pdf.embedFont(StandardFonts.Helvetica);
    for (let i = 1; i <= 3; i++) {
      const page = pdf.addPage([595, 842]);
      page.drawText(title, { x: 48, y: 760, size: 24, font, color: rgb(0, 0.35, 0.65) });
      page.drawText(`Contoso Ltd - test document. Page ${i} of 3.`, {
        x: 48,
        y: 710,
        size: 12,
        font,
      });
      page.drawText('For demonstration only. No legal effect.', { x: 48, y: 680, size: 12, font });
    }
    write(title + '.pdf', await pdf.save());
  }
  const site = Buffer.from(
    '<svg width="960" height="600" xmlns="http://www.w3.org/2000/svg"><rect width="960" height="600" fill="#e4f0f6"/><rect y="430" width="960" height="170" fill="#98b78a"/><path d="M100 430V180H380V430 M430 430V110H750V430" fill="#859dab" stroke="#536c7b" stroke-width="8"/><path d="M465 150H715 M465 210H715 M465 270H715 M465 330H715 M140 230H340 M140 290H340 M140 350H340" stroke="#d8ecf3" stroke-width="25"/></svg>',
  );
  write('Site Photo.jpg', await sharp(site).jpeg({ quality: 90 }).toBuffer());
  write(
    'Logo.png',
    await sharp(
      Buffer.from(
        '<svg width="320" height="240" xmlns="http://www.w3.org/2000/svg"><rect width="320" height="240" rx="18" fill="#0078d4"/><path d="M205 65 C100 20 100 220 205 175" fill="none" stroke="white" stroke-width="22"/></svg>',
      ),
    )
      .png()
      .toBuffer(),
  );
  // Use an installed ffmpeg first; no runtime external content is needed.
  execFileSync(
    process.env.DMS_FFMPEG_PATH || 'ffmpeg',
    [
      '-hide_banner',
      '-loglevel',
      'error',
      '-y',
      '-f',
      'lavfi',
      '-i',
      'testsrc=size=320x180:rate=15',
      '-t',
      '2',
      '-c:v',
      'libx264',
      '-pix_fmt',
      'yuv420p',
      '-movflags',
      '+faststart',
      path.join(seedRoot, 'Kickoff Recording.mp4'),
    ],
    { stdio: 'pipe' },
  );
  write(
    'Meeting Notes.txt',
    'Contoso kickoff\n\nAgenda: scope, delivery dates, commercial review.\nOwner: Alex Wilber.\n',
  );
  write(
    'Requirements.md',
    '# Contoso requirements\n\n- Accessible document library\n- Note uploads and SharePoint links\n- Metadata and audit visibility\n',
  );
  write('Export.csv', 'Item,Quantity,Price\nImplementation,1,12500\nSupport,12,450\n');
  write(
    'config.json',
    JSON.stringify({ environment: 'mock', organisation: 'Contoso', version: 1 }, null, 2),
  );
  const zip = new JSZip();
  zip.file('README.txt', 'DMS Grid sample archive.');
  zip.file('Export.csv', fs.readFileSync(path.join(seedRoot, 'Export.csv')));
  write('Archive.zip', await zip.generateAsync({ type: 'nodebuffer' }));
  write(
    'Network Diagram.vsdx',
    'DMS Grid Visio icon/unsupported-preview placeholder. This is deliberately not a valid Visio drawing.',
  );
  const snapshot: Snapshot = {
    documents: [],
    annotations: [],
    migrations: { [records.legacy]: { state: 'NotStarted', processed: 0, total: 40 } },
    binarySeeds: {},
  };
  function note(
    group: number,
    index: number,
    parent: string,
    entity: string,
    fixtureIndex: number,
  ): Entity {
    const [filename, mimetype] = fixtures[fixtureIndex % fixtures.length],
      id = seedGuid(group, index),
      data = fs.readFileSync(path.join(seedRoot, filename));
    const row: Entity = {
      annotationid: id,
      subject: filename,
      filename,
      mimetype,
      isdocument: true,
      filesize: data.length,
      _objectid_value: parent,
      [`objectid_${entity}@odata.bind`]: `/${entity === 'account' ? 'accounts' : 'opportunities'}(${parent})`,
      createdon: new Date(now - index * 3600000).toISOString(),
      modifiedon: new Date(now - index * 3600000).toISOString(),
      _createdby_value: userIds[index % 6],
      _modifiedby_value: userIds[index % 6],
      _ownerid_value: userIds[index % 6],
      '@odata.etag': 'W/"1"',
    };
    snapshot.annotations.push(row);
    snapshot.binarySeeds![id] = filename;
    return row;
  }
  function document(
    group: number,
    index: number,
    parent: string,
    fixtureIndex: number,
    link = false,
  ): Entity {
    const [filename, mimetype] = fixtures[fixtureIndex % fixtures.length],
      annotation = link ? null : note(group + 100, index, parent, 'account', fixtureIndex),
      data = fs.readFileSync(path.join(seedRoot, filename));
    const intervals = [2 * 60000, 3 * 3600000, 25 * 3600000, 8 * 86400000, 400 * 86400000];
    const date = new Date(now - intervals[index % 5] - index * 1000).toISOString();
    const row: Entity = {
      dms_documentid: seedGuid(group, index),
      dms_name: filename,
      dms_originalfilename: filename,
      dms_description:
        index === 2
          ? 'A detailed sample document description. '.repeat(35)
          : 'Sample document for local testing.',
      dms_regardingid: parent,
      dms_regardingtype: 'account',
      dms_provider: link ? 100000001 : 100000000,
      dms_storageref: link
        ? `https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Contoso/${encodeURIComponent(filename)}?web=1`
        : annotation!.annotationid,
      dms_contenttype: mimetype,
      dms_filesizekb: link ? null : Math.ceil(data.length / 1024),
      dms_checksum: link ? null : createHash('sha256').update(data).digest('hex'),
      dms_documenttype: 100000000 + (index % 6),
      dms_documentstatus: 100000000 + (index % 3),
      dms_expirydate: new Date(now + [-10, 20, 200][index % 3] * 86400000)
        .toISOString()
        .slice(0, 10),
      dms_uploadstate: 100000001,
      createdon: date,
      modifiedon: date,
      _createdby_value: userIds[index % 6],
      _modifiedby_value: userIds[index % 6],
      _ownerid_value: userIds[index % 6],
      '@odata.etag': 'W/"1"',
    };
    snapshot.documents.push(row);
    return row;
  }
  fixtures.forEach((_, index) => document(10, index + 1, records.contoso, index));
  const linkNames = [
    'Statement of Work.docx',
    'Budget 2026.xlsx',
    'Board Deck.pptx',
    'Team Notebook',
    'Sales documents',
  ];
  linkNames.forEach((name, index) => {
    const row = document(10, 15 + index, records.contoso, index, true);
    row.dms_name = name;
    row.dms_originalfilename = name;
    row.dms_storageref = `https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Contoso/${index === 4 ? '' : encodeURIComponent(name)}?web=1`;
    if (index >= 3) row.dms_contenttype = 'application/x-sharepoint-link';
  });
  for (const [index, name, state] of [
    [20, 'Failed Upload.pdf', 100000002],
    [21, 'Pending Upload.pdf', 100000000],
    [22, 'Archived Blob.pdf', 100000001],
  ] as const) {
    const row = document(10, index, records.contoso, 3);
    const id = String(row.dms_storageref);
    snapshot.annotations = snapshot.annotations.filter((note) => note.annotationid !== id);
    delete snapshot.binarySeeds![id];
    row.dms_name = name;
    row.dms_uploadstate = state;
    row.dms_storageref = null;
    if (index === 22) row.dms_provider = 100000002;
  }
  snapshot.documents[0].dms_name =
    'Master Services Agreement - Contoso regional services delivery and support schedule with detailed contractual obligations and renewal provisions 2026.docx';
  for (let i = 1; i <= 1200; i++) {
    const row = document(20, i, records.fabrikam, i % fixtures.length, i % 4 === 0);
    row.dms_name = `Fabrikam ${String(i).padStart(4, '0')} ${row.dms_name}`;
  }
  for (let i = 1; i <= 40; i++) note(130, i, records.legacy, 'opportunity', i % fixtures.length);
  write('state.json', JSON.stringify(snapshot));
  write(
    'README.txt',
    'Generated local sample files. Office files, PDFs, JPEG, PNG, MP4 and ZIP are valid/openable. Network Diagram.vsdx is an intentional preview-fallback placeholder. Images are original generated fixtures.\n',
  );
  return snapshot;
}
if (require.main === module) {
  void generateSeed()
    .then(() => {
      new Store(path.join(root, 'data'), folder).reset();
    })
    .catch((error) => {
      process.stderr.write(error instanceof Error ? error.message : 'Seed generation failed.');
      process.exitCode = 1;
    });
}
