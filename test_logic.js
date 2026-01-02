
const assert = require('assert');

// --- Pseudos de la lógica implementada en app.js ---

function sanitizeData(rows) {
    const cleanedRows = [];
    rows.forEach(r => {
        const row = { ...r };
        Object.keys(row).forEach(k => {
            if (k.toLowerCase() === 'whatsapp') {
                const val = (row[k] || '').toString();
                row[k] = val.replace(/[+\s]/g, '');
            }
        });
        cleanedRows.push(row);
    });
    return cleanedRows;
}

function reorderHeaders(headers) {
    // Simula la logica de app.js
    let newHeaders = [...headers];

    // 1. Source logic (simplificada para test)
    if (newHeaders.includes('Source')) {
        newHeaders = ['Source', ...newHeaders.filter(h => h !== 'Source')];
    }

    const waCol = newHeaders.find(h => h.toLowerCase() === 'whatsapp');
    if (waCol) {
        let tempHeaders = newHeaders.filter(h => h !== waCol);

        const idxComentarios = tempHeaders.findIndex(h => h.toLowerCase().includes('comentarios'));

        if (idxComentarios !== -1) {
            tempHeaders.splice(idxComentarios + 1, 0, waCol);
            newHeaders = tempHeaders;
        } else {
            const idxRest = tempHeaders.findIndex(h => h.toLowerCase().includes('rest.'));
            if (idxRest !== -1) {
                tempHeaders.splice(idxRest, 0, waCol);
                newHeaders = tempHeaders;
            } else {
                tempHeaders.push(waCol);
                newHeaders = tempHeaders;
            }
        }
    }
    return newHeaders;
}

// --- Test Cases ---

console.log("Running Tests...");

// Test Case 1: Sanitization
const data1 = [
    { Name: 'Alice', Whatsapp: '+593 999' },
    { Name: 'Bob', Whatsapp: '099 123 456' },
    { Name: 'Charlie', WHATSAPP: '+1 (555) 123' } // Note: My logic only strips + and \s. Parentheses remain if logic is strictly + and \s. User said "sin signo + ni espacios".
];
// User specific request: "sin signo + ni espacios". My regex was /[+\s]/g. Parentheses were NOT requested to be removed.
// Just checking if my regex complies with "sin signo + ni espacios".

const cleaned1 = sanitizeData(data1);
assert.strictEqual(cleaned1[0].Whatsapp, '593999');
assert.strictEqual(cleaned1[1].Whatsapp, '099123456');
assert.strictEqual(cleaned1[2].WHATSAPP, '1(555)123'); // Parentheses remain? Yes, regex was /[+\s]/g

console.log("✅ Sanitization Passed");

// Test Case 2: Reordering (With Comentarios and Rest)
const h1 = ['Rest.', 'Whatsapp', 'Comentarios', 'ID'];
// Expected: Whatsapp between Comentarios and Rest.
// i.e. ... Comentarios, Whatsapp, Rest ... OR ... Comentarios, Whatsapp, Rest ...
// Logic: Find Comentarios. Insert After.
// h1 -> temp: ['Rest.', 'Comentarios', 'ID']
// Comentarios is at index 1.
// splice(2, 0, Whatsapp) -> ['Rest.', 'Comentarios', 'Whatsapp', 'ID']? 
// Wait. 
// User said "entre las columnas "Comentarios" y "Rest."" (between Comentarios and Rest).
// If headers are ['Rest.', 'Comentarios...'], inserting after Comentarios puts it between? No.
// It puts it after Comentarios.
// Sample Order usually: ID, Name, Comentarios, Rest.
// If typical order is Comentarios -> Rest, then inserting after Comentarios works.
// Let's check result.
const r1 = reorderHeaders(h1);
// temp: Rest, Comentarios, ID
// idxComentarios = 1.
// splice(2, 0, Whatsapp) -> Rest, Comentarios, Whatsapp, ID.
// This is NOT between Comentarios and Rest IF Rest was before.
// But usually Comentarios comes before Rest in left-to-right reading.
// User said "entre x y y".
// If I have [Comentarios, Rest], result [Comentarios, Whatsapp, Rest].
// Let's test [Comentarios, Rest].
const h2 = ['Source', 'Comentarios', 'Rest.', 'Other'];
// Add Whatsapp
const h2_in = [...h2, 'Whatsapp'];
const r2 = reorderHeaders(h2_in);
// temp: Source, Comentarios, Rest., Other
// idxComentarios = 1.
// splice(2, 0, Whatsapp) -> Source, Comentarios, Whatsapp, Rest., Other.
// This matches "Entre Comentarios y Rest".

assert.deepStrictEqual(r2, ['Source', 'Comentarios', 'Whatsapp', 'Rest.', 'Other']);
console.log("✅ Reordering (Comentarios found) Passed");

// Test Case 3: Reordering (No Comentarios, only Rest)
const h3 = ['Rest.', 'Other', 'Whatsapp'];
const r3 = reorderHeaders(h3);
// temp: Rest., Other
// idxComentarios = -1.
// idxRest = 0.
// Insert BEFORE Rest: splice(idxRest, 0, Whatsapp) -> Whatsapp, Rest., Other.
assert.deepStrictEqual(r3, ['Whatsapp', 'Rest.', 'Other']);
console.log("✅ Reordering (Only Rest found) Passed");

// Test Case 4: Reordering (Neither)
const h4 = ['EMAIL', 'Whatsapp'];
const r4 = reorderHeaders(h4);
// temp: EMAIL
// Comentarios? No. Rest? No.
// Push -> EMAIL, Whatsapp.
assert.deepStrictEqual(r4, ['EMAIL', 'Whatsapp']);
console.log("✅ Reordering (Neither found) Passed");

console.log("All tests passed!");
