const axios = require('axios');
const cheerio = require('cheerio');
const Excel = require('exceljs');

// 設定要爬取的網址及 class 名稱，可以自行修改
const url = 'https://www.idroc.org.tw/member.php?page=';
const className = '.large-8';

async function getData(page) {
  try {
    const response = await axios.get(url + page);
    const $ = cheerio.load(response.data);
    const data = [];

    $(className).each(function () {
      const lines = $(this).text().trim().split('\n');
      const item = {
        jobTitle: '',
        name: '',
        address: '',
        phone: '',
        phone2: '',
        email: '',
        company: '',
      };

      lines.forEach((line) => {
        if (line.includes('會員代表') || line.includes('學會理事')) {
          item.jobTitle = line.trim();
        } else if (line.includes('-')) {
          item.phone += line.trim() + '\n';
          item.phone2 += line.trim() + '\n';
        } else if (line.includes('@')) {
          item.email = line.trim();
        } else if (
          line.includes('區') ||
          line.includes('市') ||
          line.includes('縣')
        ) {
          item.address = line.trim();
        } else {
          if (!item.name) {
            item.name = line.trim();
          } else {
            item.company = line.trim();
          }
        }
      });

      item.phone = item.phone.replace(/\n$/, '');
      item.phone2 = item.phone2.replace(/\n$/, '');

      data.push(item);
    });

    return data;
  } catch (error) {
    console.log(error);
  }
}

async function crawlPages(start, end) {
  try {
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    worksheet.columns = [
      { header: 'Job Title', key: 'jobTitle' },
      { header: 'Name', key: 'name' },
      { header: 'Address', key: 'address' },
      { header: 'Phone', key: 'phone' },
      { header: 'Phone2', key: 'phone2' },
      { header: 'Email', key: 'email' },
      { header: 'Company', key: 'company' },
    ];

    for (let i = start; i <= end; i++) {
      const data = await getData(i);
      data.forEach((item) => {
        const rows = item.phone.split('\n') || item.phone2.split('\n');
        rows.forEach((row, index) => {
          worksheet.addRow({
            jobTitle: item.jobTitle,
            name: index === 0 ? item.name : '',
            address: index === 0 ? item.address : '',
            phone: row.trim(),
            phone2: row.trim(),
            email: index === 0 ? item.email : '',
            company: index === 0 ? item.company : '',
          });
        });
      });
    }

    await workbook.xlsx.writeFile('output.xlsx');
    console.log('Data saved to output.xlsx');
  } catch (error) {
    console.log(error);
  }
}

crawlPages(1, 18); // 爬取第1頁到第18頁的資料
