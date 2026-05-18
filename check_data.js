const XLSX = require('xlsx');
const wb = XLSX.readFile('data_magang.xlsx');
const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

const ds = data.filter(r => {
    const m = String(r['Major'] || r['Major '] || '').toLowerCase();
    return m.includes('data science');
});

console.log('DS total:', ds.length);
ds.forEach((r, i) => {
    const name = String(r['Student Name'] || r['Student Name '] || '').trim();
    const company = String(r['Company Name'] || r['Company Name '] || '').trim();
    const position = String(r['Position'] || r['Position '] || '').trim();
    const sem = String(r['Internship Semester'] || r['Internship Semester '] || '').trim();
    const comp = String(r['Compensation'] || r['Compensation '] || '').trim();
    const type = String(r['Type'] || r['Type '] || '').trim();
    const ev = String(r['Evidence'] || r['Evidence '] || '').trim();
    const fb = String(r["Employer's Feedback"] || r["Employer's Feedback "] || '').trim();
    
    console.log(`${i+1} | ${name} | ${company} | ${position} | Sem:${sem} | Comp:${comp} | Type:${type} | Ev:${ev ? 'YES' : '-'} | Fb:${fb ? 'YES' : '-'}`);
});
