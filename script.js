import jsdom from 'jsdom';
import XslxPopulate from 'xlsx-populate';

async function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getLinks(pageInit, pageEnd) {
  const links = [];
  try {
    for (let i = pageInit; i <= pageEnd; i++) {
      const url = `https://guiaempresas.universia.net.co/localidad/YUMBO/?qPagina=${i}`;
      const response = await fetch(url, {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
        },
      });
      const data = await response.text();
      const doc = new jsdom.JSDOM(data).window.document;
      const anchors = doc.querySelectorAll('td.textleft a');
      anchors.forEach((anchor) => {
        links.push(anchor.href);
      });

      // Introduce un retraso de 1 segundo entre cada solicitud
      await delay(1000); // 1000 milisegundos = 1 segundo
    }
  } catch (error) {
    console.log(error);
  }
  return links;
}

async function getInfoCompany(url) {
  const res = {};

  const dataRequired = ['Dirección:', 'Nit:', 'Teléfono:', 'Actividad:'];
  try {
    const response = await fetch(
      `https://guiaempresas.universia.net.co${url}`,
      {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
        },
      }
    );
    const data = await response.text();
    const doc = new jsdom.JSDOM(data).window.document;
    res.title = doc.querySelector('h1').textContent;
    const rows = doc.querySelectorAll('div#ficha_iden tr');

    rows.forEach((row) => {
      const th = row.querySelector('th');
      const td = row.querySelector('td');

      const nombreCampo = th.textContent.trim();
      const valorCampo = td.textContent;

      if (dataRequired.includes(nombreCampo)) {
        res[nombreCampo] = valorCampo;
      }
    });

    // Introduce un retraso de 1 segundo después de cada solicitud de información de empresa
    await delay(1000); // 1000 milisegundos = 1 segundo
  } catch (error) {
    console.log(error);
  }
  console.log(res);
  return res;
}

async function main() {
  const resultExcel = await XslxPopulate.fromBlankAsync('companies.xlsx');

  const sheet = resultExcel.sheet(0);
  sheet.cell('A1').value('Nombre');
  sheet.cell('B1').value('Dirección');
  sheet.cell('C1').value('Nit');
  sheet.cell('D1').value('Teléfono');
  sheet.cell('E1').value('Actividad');

  const links = await getLinks(45, 46);
  console.log(links);

  const companies = [];
  for (const link of links) {
    const company = await getInfoCompany(link);
    companies.push(company);
  }

  console.log(companies);

  companies.forEach((company, index) => {
    sheet.cell(`A${index + 2}`).value(company.title);
    sheet.cell(`B${index + 2}`).value(company['Dirección:']);
    sheet.cell(`C${index + 2}`).value(company['Nit:']);
    sheet.cell(`D${index + 2}`).value(company['Teléfono:']);
    sheet.cell(`E${index + 2}`).value(company['Actividad:']);
  });

  await resultExcel.toFileAsync('companies.xlsx');
}

main();
