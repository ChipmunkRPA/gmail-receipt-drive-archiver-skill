const { chromium } = require('playwright-core');

async function main() {
  const [htmlPath, pdfPath, chromePath] = process.argv.slice(2);
  if (!htmlPath || !pdfPath || !chromePath) {
    throw new Error('Usage: node render_email_html_to_pdf.js <htmlPath> <pdfPath> <chromePath>');
  }

  const browser = await chromium.launch({
    executablePath: chromePath,
    headless: true,
  });

  try {
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(15000);
    const fileUrl = new URL(`file:///${htmlPath.replace(/\\/g, '/')}`);
    try {
      await page.goto(fileUrl.toString(), { waitUntil: 'domcontentloaded', timeout: 15000 });
    } catch (error) {
      // Some emails reference remote assets; continue and print what rendered.
    }
    await page.emulateMedia({ media: 'screen' });
    await page.waitForTimeout(4000);
    await page.pdf({
      path: pdfPath,
      format: 'Letter',
      printBackground: true,
      margin: {
        top: '0.5in',
        right: '0.5in',
        bottom: '0.5in',
        left: '0.5in',
      },
    });
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
