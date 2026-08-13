import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const outDir='ub-international-dashboard/data-source';
const previewDir='ub-international-dashboard/qa';
const headers=['Activity ID','Activity Title','Activity Description','Activity Category','Activity Status','Start Date','End Date','Year','Agreement Availability','Agreement Reference Type','MoU ID','MoU Name','Agreement Number','Leading UB Faculty','Involved UB Faculties','Involved UB Study Programs','Responsible UB Unit','UB Person in Charge','UB PIC Email','Partner Institution','Partner Country','Partner Contact Person','Partner Contact Email','Activity Mode','Activity Location','Mobility Direction','International Participant Count','International Participant Role','UB Participant Count','UB Participant Role','Funding Source','Funding Amount','Currency','Expected Output','Actual Output','Output Type','Output Link','Evidence Link','Related SDGs','Remarks','Last Updated'];
const categories=['Student Mobility','Joint Research','Visiting Lecturer','Joint Publication','Joint Supervision','Guest Lecture','Double Degree','Joint Degree','Staff Mobility','Joint International Scientific Event','Joint Community Service','Internship','Summer School','Winter School','COIL','Visiting Researcher','Adjunct Professor','Joint Workshop','Joint Seminar','Joint Conference','Research Grant','Staff Training','Curriculum Development','Academic Benchmarking','Laboratory Collaboration','Innovation or Startup Collaboration','Student Competition','Cultural Exchange','Other'];
const wb=Workbook.create();
const data=wb.worksheets.add('Activities');
const lookups=wb.worksheets.add('Lookups');
const guide=wb.worksheets.add('Instructions');
data.getRange('A1:AO2').values=[headers,Array(headers.length).fill('')];
data.getRange('A1:AO1').format={fill:'#003B70',font:{bold:true,color:'#FFFFFF'},rowHeight:40,wrapText:true,verticalAlignment:'center'};
data.getRange('A2:AO2').format={fill:'#F8FAFC',font:{color:'#334155'},rowHeight:24};
data.getRange('F:G').format.numberFormat='yyyy-mm-dd'; data.getRange('AO:AO').format.numberFormat='yyyy-mm-dd';
data.getRange('H2').formulas=[['=IF(F2="","",YEAR(F2))']];
data.freezePanes.freezeRows(1); data.freezePanes.freezeColumns(2); data.showGridLines=false;
const widths=[18,36,52,28,18,14,14,10,24,24,18,42,22,32,42,42,32,28,28,38,22,28,28,16,32,20,18,26,18,24,24,18,12,36,36,22,42,42,20,42,16];
const colName=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s};
for(let i=0;i<widths.length;i++) data.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=widths[i];
data.tables.add('A1:AO2',true,'ActivitiesTable').style='TableStyleMedium2';
data.getRange('D2:D5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$A$2:$A$30"}};
data.getRange('E2:E5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$B$2:$B$6"}};
data.getRange('I2:I5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$C$2:$C$3"}};
data.getRange('J2:J5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$D$2:$D$7"}};
data.getRange('X2:X5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$E$2:$E$4"}};
data.getRange('Z2:Z5000').dataValidation={rule:{type:'list',formula1:"'Lookups'!$F$2:$F$6"}};

const lookupHeaders=['Activity Categories','Activity Status','Agreement Availability','Reference Type','Activity Mode','Mobility Direction','Participant Roles','Output Types'];
const cols=[categories,['Planned','Ongoing','Completed','Postponed','Cancelled'],['With Agreement','Without Agreement'],['MoU','MoA','IA','LoI','Letter of Cooperation','None'],['Online','Offline','Hybrid'],['Incoming','Outgoing','Reciprocal','Not Applicable','Mixed'],['Student','Lecturer','Professor','Researcher','Staff','Practitioner','Mixed'],['Publication','Mobility','Research Grant','Curriculum','Event','Patent','Product','Report','Other']];
lookups.getRange('A1:H31').values=Array.from({length:31},(_,r)=>Array.from({length:8},(_,c)=>r===0?lookupHeaders[c]:(cols[c][r-1]??'')));
lookups.getRange('A1:H1').format={fill:'#D4A017',font:{bold:true,color:'#FFFFFF'},rowHeight:28};
lookups.getRange('A:H').format.columnWidth=28; lookups.freezePanes.freezeRows(1); lookups.showGridLines=false;

const instructions=[
 ['UB International Activity Database Template'],
 ['Purpose','Enter one international collaboration activity per row. This flat table feeds the static dashboard.'],
 ['Required minimum','Activity ID, Activity Title, Activity Category, Activity Status, Start Date, Agreement Availability, Partner Institution, Partner Country.'],
 ['Agreement rule','With Agreement: enter a valid MoU ID when applicable. Without Agreement: MoU ID, MoU Name, and Agreement Number may remain blank.'],
 ['Multiple values','Separate multiple faculties, study programs, participant roles, and SDGs using | (vertical bar).'],
 ['Dates','Use real Excel dates; display format is yyyy-mm-dd. Year is calculated from Start Date.'],
 ['Identifiers','Use a stable unique Activity ID, e.g. ACT-2026-001. Do not reuse IDs.'],
 ['Privacy','Do not publish sensitive personal information in data intended for a public dashboard.'],
 ['Update workflow','Update this workbook → run python scripts/prepare_data.py → run npm run build → deploy dist/.'],
];
guide.getRange('A1:B9').values=instructions.map(r=>r.length===1?[r[0],'']:r);
guide.getRange('A1:B1').merge(); guide.getRange('A1').format={fill:'#003B70',font:{bold:true,color:'#FFFFFF',size:16},rowHeight:34};
guide.getRange('A2:A9').format={fill:'#E8F0F8',font:{bold:true,color:'#003B70'},wrapText:true}; guide.getRange('B2:B9').format={wrapText:true};
guide.getRange('A:A').format.columnWidth=24; guide.getRange('B:B').format.columnWidth=90; guide.getRange('A2:B9').format.rowHeight=44; guide.showGridLines=false;

await fs.mkdir(outDir,{recursive:true}); await fs.mkdir(previewDir,{recursive:true});
for(const name of ['Activities','Lookups','Instructions']){const range=name==='Activities'?'A1:AO8':undefined;const img=await wb.render({sheetName:name,range,autoCrop:range?undefined:'all',scale:1,format:'png'});await fs.writeFile(`${previewDir}/activity_template_${name}.png`,new Uint8Array(await img.arrayBuffer()));}
const file=await SpreadsheetFile.exportXlsx(wb); await file.save(`${outDir}/Database_Aktivitas_Kerjasama_Internasional.xlsx`);
console.log((await wb.inspect({kind:'workbook,sheet,table',maxChars:6000,tableMaxRows:5,tableMaxCols:8})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',options:{useRegex:true,maxResults:100}})).ndjson);
