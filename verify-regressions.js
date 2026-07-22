const { execFileSync } = require('node:child_process');
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
      if (entry !== '.' && entry !== '..') x2t.FS.unlink(`${path}/${entry}`);
    }
    return;
  }
  x2t.FS.mkdir(path);
}

function convertExampleDocumentTitle() {
  recreateDirectory('/working');
  const inputName = 'example-document-title.odt';
  const outputName = 'example-document-title.docx';
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

function verifyExampleDocumentTitleChart() {
  const chartXml = readZipEntry(
    'results/example-document-title.docx',
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

  const message = 'Verified Example Document Title chart gradient colors.\n';
  fs.writeFileSync('results/verify-regressions.js.log', message);
  process.stdout.write(message);
}

x2t.onRuntimeInitialized = function () {
  try {
    convertExampleDocumentTitle();
    verifyExampleDocumentTitleChart();
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  }
};
