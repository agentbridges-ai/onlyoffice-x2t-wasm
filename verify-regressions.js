const { execFileSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');

const x2t = require('./x2t');

function readZipEntry(archivePath, entryPath) {
  return execFileSync('unzip', ['-p', archivePath, entryPath], {
    encoding: 'utf8',
    maxBuffer: 4 * 1024 * 1024,
  });
}

function recreateDirectory(path) {
  if (x2t.FS.analyzePath(path).exists) {
    for (const entry of x2t.FS.readdir(path)) {
      if (entry === '.' || entry === '..') continue;
      const child = `${path}/${entry}`;
      const stat = x2t.FS.stat(child);
      if (x2t.FS.isDir(stat.mode)) {
        recreateDirectory(child);
        x2t.FS.rmdir(child);
      } else {
        x2t.FS.unlink(child);
      }
    }
    return;
  }
  x2t.FS.mkdir(path);
}

function sha256(data) {
  return crypto.createHash('sha256').update(data).digest('hex');
}

function ensureDirectory(path) {
  const parts = path.split('/').filter(Boolean);
  let current = '';
  for (const part of parts) {
    current += `/${part}`;
    if (!x2t.FS.analyzePath(current).exists) x2t.FS.mkdir(current);
  }
}

function runCanvasRegression({
  inputName,
  inputSha256,
  outputSha256,
  outputSize,
  formatFrom,
  formatTo,
  header,
  expectedMedia,
}) {
  const input = fs.readFileSync(`tests/${inputName}`);
  if (sha256(input) !== inputSha256) {
    throw new Error(`${inputName} regression fixture digest changed`);
  }

  const convert = () => {
    recreateDirectory('/working');
    ensureDirectory('/working/media');
    ensureDirectory('/working/themes');
    ensureDirectory('/working/fonts');
    ensureDirectory('/tmp/x2t-conversion');
    x2t.FS.writeFile(`/working/${inputName}`, input);
    x2t.FS.writeFile(
      '/working/params.xml',
      `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>/working/${inputName}</m_sFileFrom>
  <m_sFileTo>/working/Editor.bin</m_sFileTo>
  <m_sThemeDir>/working/themes</m_sThemeDir>
  <m_sFontDir>/working/fonts/</m_sFontDir>
  <m_sTempDir>/tmp/x2t-conversion</m_sTempDir>
  <m_nFormatFrom>${formatFrom}</m_nFormatFrom>
  <m_nFormatTo>${formatTo}</m_nFormatTo>
  <m_bIsNoBase64>false</m_bIsNoBase64>
</TaskQueueDataConvert>`,
    );
    const result = x2t.ccall('main1', 'number', ['string'], ['/working/params.xml']);
    if (result !== 0) {
      throw new Error(`${inputName} Canvas conversion failed with exit code ${result}`);
    }
    return Buffer.from(x2t.FS.readFile('/working/Editor.bin'));
  };

  // Release reproducibility is checked from a cold x2t module. The converter
  // retains document-local caches between main1 calls, so a warm second call
  // is not a byte-for-byte clean-build comparison. Every CI invocation loads
  // this verifier in a fresh Node process, while two independent no-cache
  // workflow runs compare the resulting golden and release artifacts.
  const output = convert();
  const actualHeader = Buffer.from(output.subarray(0, header.length)).toString('ascii');
  if (actualHeader !== header || output.length !== outputSize || sha256(output) !== outputSha256) {
    throw new Error(
      `${inputName} Canvas output changed: ${actualHeader}, ${output.length} bytes, ${sha256(output)}`,
    );
  }
  const media = x2t.FS.readdir('/working/media').filter((name) => name !== '.' && name !== '..').sort();
  for (const name of expectedMedia) {
    if (!media.includes(name)) throw new Error(`${inputName} lost media/${name}`);
  }
  process.stdout.write(
    `Verified ${inputName} -> ${header} (${outputSize} bytes, sha256 ${outputSha256}).\n`,
  );
}

const canvasRegressions = {
  'example-document-title-ole.doc': {
    inputName: 'example-document-title-ole.doc',
    inputSha256: 'd85e44ae5368ccbbe57ded8533ced05a250c30cfa15da10f19fdaf63f080238c',
    outputSha256: '074a9b350ff6a6e1ee32866c03416a0682c05635cdb8f3f60b6e4a02eaad9a2a',
    outputSize: 132030,
    formatFrom: 66,
    formatTo: 8193,
    header: 'DOCY;v5;',
    expectedMedia: ['display6image1.bin', 'display6image1.emf', 'display6image1.svg'],
  },
  'pivot-slicer-showcase.xlsx': {
    inputName: 'pivot-slicer-showcase.xlsx',
    inputSha256: 'ffecc0a33c9e41b392fbee30127a97f3e5c3577c717be103471460bd07c2ec58',
    outputSha256: 'c40fb3f4f67311426110d4786eb4684981aec9ed05b4f13c6367c8470de4d89e',
    outputSize: 85138,
    formatFrom: 257,
    formatTo: 8194,
    header: 'XLSY;v2;',
    expectedMedia: ['image1.png'],
  },
};

function verifyNativeOfficeCanvasModels() {
  const workbookXml = readZipEntry(
    'tests/pivot-slicer-showcase.xlsx',
    'xl/workbook.xml',
  );
  const coreXml = readZipEntry(
    'tests/pivot-slicer-showcase.xlsx',
    'docProps/core.xml',
  );
  if (/x15ac:absPath|\/Users\//.test(workbookXml) || />Y X</.test(coreXml)) {
    throw new Error('Pivot/Slicer fixture contains local author or filesystem metadata');
  }

  for (const inputName of Object.keys(canvasRegressions)) {
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      const output = execFileSync(
        process.execPath,
        [__filename, '--canvas-regression', inputName],
        { encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 },
      );
      process.stdout.write(`Cold-module attempt ${attempt}: ${output}`);
    }
  }
}

function convertExampleTitleOdt() {
  recreateDirectory('/working');
  const inputName = 'example-title.odt';
  const outputName = 'example-title.docx';
  const input = fs.readFileSync(`tests/${inputName}`);
  x2t.FS.writeFile(`/working/${inputName}`, input);
  x2t.FS.writeFile(
    '/working/params.xml',
    `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>/working/${inputName}</m_sFileFrom>
  <m_sFileTo>/working/${outputName}</m_sFileTo>
  <m_nFormatFrom>67</m_nFormatFrom>
  <m_nFormatTo>65</m_nFormatTo>
  <m_bIsNoBase64>false</m_bIsNoBase64>
</TaskQueueDataConvert>`,
  );
  const result = x2t.ccall('main1', 'number', ['string'], ['/working/params.xml']);
  if (result !== 0) throw new Error(`ODT to DOCX regression conversion failed with exit code ${result}`);
  fs.writeFileSync(`results/${outputName}`, x2t.FS.readFile(`/working/${outputName}`));
}

function verifyExampleTitleChart() {
  const chartXml = readZipEntry(
    'results/example-title.docx',
    'word/charts/chart1.xml',
  );
  const expectedGradientColors = [
    ['3574AC', '4697E0', '4397E4'],
    ['C5590F', 'FF7415', 'FF7416'],
    ['7B7B7B', '9F9F9F', 'A0A0A0'],
    ['BE8F00', 'F7BA00', 'F8BA00'],
    ['2451A0', '2E69D0', '2C68D4'],
  ];

  const gradientFills = chartXml.match(/<a:gradFill\b[\s\S]*?<\/a:gradFill>/g) || [];
  if (gradientFills.length !== expectedGradientColors.length) {
    throw new Error(
      `Expected ${expectedGradientColors.length} chart gradient fills, found ${gradientFills.length}`,
    );
  }

  gradientFills.forEach((gradientFill, index) => {
    const colors = Array.from(
      gradientFill.matchAll(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/g),
      (match) => match[1].toUpperCase(),
    );
    const expected = expectedGradientColors[index];
    if (colors.length !== expected.length || colors.some((color, colorIndex) => color !== expected[colorIndex])) {
      throw new Error(
        `Gradient ${index + 1} expected ${expected.join(', ')}; found ${colors.join(', ') || 'no colors'}`,
      );
    }
  });

  const message = 'Verified Example Title chart gradient colors.\n';
  fs.writeFileSync('results/verify-regressions.js.log', message);
  process.stdout.write(message);
}

function verifyExampleTitleOdpImport() {
  recreateDirectory('/odp-working');
  x2t.FS.mkdir('/odp-working/tmp');
  x2t.FS.writeFile('/odp-working/example-title.odp', fs.readFileSync('tests/example-title.odp'));
  x2t.FS.writeFile(
    '/odp-working/params.xml',
    `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>/odp-working/example-title.odp</m_sFileFrom>
  <m_sTempDir>/odp-working/tmp</m_sTempDir>
  <m_sFileTo>/odp-working/example-title.bin</m_sFileTo>
  <m_nFormatFrom>131</m_nFormatFrom>
  <m_nFormatTo>4099</m_nFormatTo>
  <m_bIsNoBase64>false</m_bIsNoBase64>
</TaskQueueDataConvert>`,
  );
  const result = x2t.ccall('main1', 'number', ['string'], ['/odp-working/params.xml']);
  if (result !== 0) throw new Error(`ODP to PPTY regression conversion failed with exit code ${result}`);
  const output = x2t.FS.readFile('/odp-working/example-title.bin');
  if (output.length < 1024) throw new Error(`ODP to PPTY regression output is unexpectedly small: ${output.length}`);
  process.stdout.write('Verified Example Title ODP imports without a WebAssembly function signature mismatch.\n');
}

function verifyFb2Export() {
  const outputPath = 'results/html-export.fb2';
  const output = fs.readFileSync(outputPath, 'utf8');
  if (!/<FictionBook\b/i.test(output)) {
    throw new Error(`FB2 regression output is not a FictionBook document: ${outputPath}`);
  }
  process.stdout.write('Verified HTML exports to a FictionBook document.\n');
}

function verifyHtmlDerivedExports() {
  const markdown = fs.readFileSync('results/html-export.md', 'utf8');
  if (!markdown.includes('This paragraph must survive')) {
    throw new Error('HTML to Markdown regression output lost the document paragraph.');
  }

  const epubMimeType = readZipEntry('results/html-export.epub', 'mimetype').trim();
  if (epubMimeType !== 'application/epub+zip') {
    throw new Error(`HTML to EPUB regression output has an invalid mimetype: ${epubMimeType}`);
  }
  process.stdout.write('Verified HTML exports to EPUB and Markdown documents.\n');
}

x2t.onRuntimeInitialized = function () {
  try {
    if (process.argv[2] === '--canvas-regression') {
      const regression = canvasRegressions[process.argv[3]];
      if (!regression) {
        throw new Error(`Unknown Canvas regression fixture: ${process.argv[3] || '(empty)'}`);
      }
      runCanvasRegression(regression);
      return;
    }

    convertExampleTitleOdt();
    verifyExampleTitleChart();
    verifyExampleTitleOdpImport();
    verifyNativeOfficeCanvasModels();
    verifyFb2Export();
    verifyHtmlDerivedExports();
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  }
};
