/* Optional offline mirror of the same pinned CDN versions, fetched with TLS verification. */
const fs = require('node:fs');
const path = require('node:path');
exports.install = async page => {
  if (!process.env.TMS_CDN_CACHE) return;
  const files = {
    '/ajax/libs/xlsx/0.18.5/xlsx.full.min.js':'xlsx.full.min.js',
    '/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js':'html2pdf.bundle.min.js',
    '/ajax/libs/Chart.js/4.4.1/chart.umd.min.js':'chart.umd.min.js'
  };
  await page.route('https://cdnjs.cloudflare.com/**', route => {
    const file = files[new URL(route.request().url()).pathname];
    if (!file) return route.continue();
    return route.fulfill({status:200,contentType:'application/javascript',body:fs.readFileSync(path.join(process.env.TMS_CDN_CACHE,file))});
  });
};
