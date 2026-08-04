const fs = require('fs');
const path = 'D:\\桌面\\工厂工作台.py';
let data = fs.readFileSync(path, 'utf8');
let unix = data.replace(/\r\n/g, '\n');

// ============================================================
// Revert workshop page: remove checkbox from items
// ============================================================
// Workshop page uses 'f_all' (not 'af_all') - use this to distinguish
// Find the workshop loadJobs section and revert its li.innerHTML

// The workshop page has: <div class="top"><input type="checkbox"... 
// We need to find the workshop version and remove the checkbox

// Workshop page context: has setFilter with f_all (without 'a')
// Find li.innerHTML in workshop page (before the admin page starts)
const wsMarker = `filter-bar"><button id="f_all"`;
const wsIdx = unix.indexOf(wsMarker);
const adminMarker = `filter-bar"><button id="af_all"`;
const adminIdx = unix.indexOf(adminMarker);

console.log('Workshop filter at:', wsIdx, 'Admin filter at:', adminIdx);

// The workshop page's loadJobs is between wsMarker and adminMarker
// Find li.innerHTML in workshop page
const searchStart = wsIdx;
const searchEnd = adminIdx;

const wsSection = unix.substring(searchStart, searchEnd);
// Find li.innerHTML with checkbox
const checkboxItem = `<input type="checkbox" class="cb-item" data-id="'+i.id+'" onchange="updateSelCount()">`;
const wsItemIdx = wsSection.indexOf(checkboxItem);

if (wsItemIdx >= 0) {
    console.log('Found checkbox in workshop page at relative offset', wsItemIdx);
    // Remove just the checkbox input from the workshop section
    const globalIdx = searchStart + wsItemIdx;
    const before = unix.substring(0, globalIdx);
    const after = unix.substring(globalIdx + checkboxItem.length);
    unix = before + after;
    console.log('Removed checkbox from workshop page');
} else {
    console.log('No checkbox found in workshop section');
}

// ============================================================
// Remove extra functions from workshop page (toggleAll, updateSelCount, deleteSelected)
// These were added between cancelJob and openComplete
// ============================================================
// Find cancelJob + extra functions + openComplete in workshop section
const wsSection2 = unix.substring(searchStart, searchEnd);
const extraFuncsStart = wsSection2.indexOf(`function toggleAll()`);
if (extraFuncsStart >= 0) {
    const globalStart = searchStart + extraFuncsStart;
    const funcEnd = wsSection2.indexOf(`function openComplete(btn)`, extraFuncsStart);
    if (funcEnd >= 0) {
        const globalEnd = searchStart + funcEnd;
        const before = unix.substring(0, globalStart);
        const after = unix.substring(globalEnd);
        unix = before + after;
        console.log('Removed extra functions from workshop page');
    }
}

data = unix.replace(/\n/g, '\r\n');
fs.writeFileSync(path, data, 'utf8');
console.log('DONE!');
